Attribute VB_Name = "ExportarExcel"
Option Explicit

Private Enum enmMenuExportacion
  mExportarPlanilla
  mExportarEstadoLineas
  mSavetoTTMDatabase
  mSavePermittingData
End Enum


'   13/01/2012
'   NO SE USA MAS, lee directo de DSinfo
'
'Public Sub Exportar(Index As Integer)
''Boton de exportar
'  On Error GoTo ErrorHandler
'
'  frmPozosFuturos.Enabled = False
'  Select Case Index
'    Case mExportarPlanilla
'      ExportarPlanilla frmPozosFuturos, frmPozosFuturos.mfg
'    Case mExportarEstadoLineas
'      ExportarEstadoLineas frmPozosFuturos, frmPozosFuturos.mfg
'    Case mSavetoTTMDatabase
'      'MsgBox "La funcionalidad Exportar a Permitting Data no esta disponible momentaneamente. El usuario oxy_read no tiene privilegios para crear la vista TTM_LOAD_V", vbInformation, "Falta de permisos en AHPSPP"
'      ExportarTTMDatabase frmPozosFuturos, frmPozosFuturos.mfg
'    Case mSavePermittingData
'      'MsgBox "La funcionalidad Exportar a Permitting Data no esta disponible momentaneamente. El usuario oxy_read no tiene privilegios para editar datos en AHPSPP (ejecutar procedure: TTM.TRUNC_TTM_PERMIT_DATA, TTM_InsertPermittingData_pkg)", vbInformation, "Falta de permisos en AHPSPP"
'      ExportarPermittingData frmPozosFuturos, frmPozosFuturos.mfg
'  End Select
'  frmPozosFuturos.Barra.Panels("info") = ""
'  frmPozosFuturos.Enabled = True
'
'ErrorHandler:
'  ErrorHandler
'End Sub


Private Sub ExportarPlanilla(Form As Form, mfg As MSHFlexGrid)
'Exporta la planilla entera tal cual se ve
    On Error Resume Next
    
    'Variables para la aplicación objExcel, el libro y la hoja
    Dim objExcel As New Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja  As Excel.Worksheet
    'Para las filas y columnas del mfg y la Hoja
    Dim Fila As Integer
    Dim columna As Integer
    Dim contColumnas As Integer
      
    ' crea los objetos y agrega el libro y la hoja
    Set Libro = objExcel.Workbooks.Add
    Set Hoja = Libro.Worksheets.Add
    Form.cmd.CancelError = True
    Form.cmd.ShowSave
    If Err.Number <> 0 Then
      GoTo ErrorHandler
    End If
    Dim colSites As Integer
    colSites = -1
    ' Recorremos el mfg por filas y columnas
    For Fila = 0 To mfg.Rows - 1
      contColumnas = 1
      Form.Barra.Panels("info") = "Generando Excel... Fila " & Fila & " de " & mfg.Rows
      For columna = 1 To mfg.Cols - 1
        'Agrega el Valor en la celda indicada del objExcel
        If mfg.ColWidth(columna) > 0 Then
          If IsDate(mfg.TextMatrix(Fila, columna)) Then
            Hoja.Columns(contColumnas).Select
            objExcel.Selection.NumberFormat = "yyyy/mm/dd"
            Hoja.Cells(Fila + 1, contColumnas).Value = Format(mfg.TextMatrix(Fila, columna), "yyyy/mm/dd")
          Else
            If Fila = 0 Then
              If mfg.TextMatrix(Fila, columna) = "Ultimo Site Visit with DMA" Then
                colSites = columna
                Hoja.Cells(Fila + 1, contColumnas).Value = "Date Visit DMA"
                contColumnas = contColumnas + 1
                Hoja.Cells(Fila + 1, contColumnas).Value = "Numero Acta Visit DMA"
                contColumnas = contColumnas + 1
                Hoja.Cells(Fila + 1, contColumnas).Value = "Comments ultimo Site Visit with DMA"
'                contColumnas = contColumnas + 1
'                Hoja.Cells(Fila + 1, contColumnas).Value = "Autor Comment Visit DMA"
              Else
                Hoja.Cells(Fila + 1, contColumnas).Value = mfg.TextMatrix(Fila, columna)
              End If
            Else
              If columna = colSites Then
              Dim strFecha As String
                
                strFecha = Trim(Mid(mfg.TextMatrix(Fila, columna), 5, 11))
                If IsDate(strFecha) Then
                  Hoja.Columns(contColumnas).Select
                  objExcel.Selection.NumberFormat = "yyyy/mm/dd"
                  Hoja.Cells(Fila + 1, contColumnas).Value = Format(strFecha, "yyyy/mm/dd")
                  contColumnas = contColumnas + 1
                  Hoja.Cells(Fila + 1, contColumnas).Value = Mid(mfg.TextMatrix(Fila, columna), 18, InStr(Mid(mfg.TextMatrix(Fila, columna), 17), " | ") - 2)
                  contColumnas = contColumnas + 1
                  Hoja.Cells(Fila + 1, contColumnas).Value = Mid(mfg.TextMatrix(Fila, columna), InStr(Mid(mfg.TextMatrix(Fila, columna), 17), " | ") + 19)
                  'contColumnas = contColumnas + 1
                  'Hoja.Cells(Fila + 1, contColumnas).Value = Right(mfg.TextMatrix(Fila, columna), InStr(StrReverse(mfg.TextMatrix(Fila, columna)), " | "))
                Else
                  contColumnas = contColumnas + 2
'                  If Len(Trim(mfg.TextMatrix(Fila, columna))) > 0 Then
'                    MsgBox mfg.TextMatrix(Fila, columna)
'                  End If
                  Hoja.Cells(Fila + 1, contColumnas).Value = mfg.TextMatrix(Fila, columna)
                End If
              Else
                Hoja.Cells(Fila + 1, contColumnas).Value = mfg.TextMatrix(Fila, columna)
              End If
            End If
          End If
          contColumnas = contColumnas + 1
        End If
        DoEvents
      Next columna
      If Not IsNumeric(mfg.TextMatrix(Fila, colIDPozo)) And Fila > 0 Then
        With objExcel
            .Rows(Fila + 1).Select
            .Selection.Font.Bold = True
            .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With .Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            .Selection.Borders(xlInsideVertical).LineStyle = xlNone
            With .Selection.Interior
                .ColorIndex = 40
                .Pattern = xlSolid
            End With
        End With
      End If
    Next Fila
    Form.Barra.Panels("info") = "Aplicando Formato..."
    AplicarMacroLista objExcel, Hoja, contColumnas
    Libro.SaveAs (Form.cmd.FileName)
    Libro.Saved = True
    objExcel.Quit

ErrorHandler:
  If Err.Number <> 0 Then
  'Cierra la hoja y el la aplicación objExcel
    If Not Libro Is Nothing Then Libro.Saved = True
    If Not objExcel Is Nothing Then: objExcel.Quit
    'Liberar los objetos
    Set objExcel = Nothing
    Set Libro = Nothing
    Set Hoja = Nothing
    Err.Clear
  End If
End Sub


Private Sub AplicarMacroLista(objExcel As Excel.Application, Hoja As Excel.Worksheet, contColumnas As Integer)
'Aplica la macro al exportar de toda la lista
  On Error Resume Next
  
  Dim columna As Long
  With objExcel
    .Rows("1:1").Select
    .Selection.Font.Bold = True
    With .Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    For columna = 1 To contColumnas
      Hoja.Columns(columna).AutoFit
    Next
    Hoja.Range("A1:BV1").Select
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Interior
        .ColorIndex = 15
        .Pattern = xlSolid
    End With
  End With
End Sub


Private Sub ExportarEstadoLineas(Form As Form, mfg As MSHFlexGrid)
'hace el reporte de estado de lineas
     On Error Resume Next
     
    'Variables para la aplicación objExcel, el libro y la hoja
    Dim objExcel As New Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja  As Excel.Worksheet
    'Para las filas y columnas del mfg y la Hoja
    Dim Fila As Integer
    Dim columna As Integer
    Dim contColumnas As Integer
      
    ' crea los objetos y agrega el libro y la hoja
    Set Libro = objExcel.Workbooks.Add
    Set Hoja = Libro.Worksheets.Add
    Form.cmd.CancelError = True
    Form.cmd.ShowSave
    If Err.Number <> 0 Then
      GoTo ErrorHandler
    End If
    
    ' Recorremos el mfg por filas y columnas
        
    For Fila = 0 To mfg.Rows - 1
      Form.Barra.Panels("info") = "Generando Excel... Fila " & Fila & " de " & mfg.Rows
      If IsNumeric(mfg.TextMatrix(Fila, colIDPozo)) Or Fila = 0 Then
        Hoja.Cells(Fila + 1, 1).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colUbicacion)), mfg.TextMatrix(Fila, colUbicacion))
        Hoja.Cells(Fila + 1, 2).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colWellID)), mfg.TextMatrix(Fila, colWellID))
        Hoja.Cells(Fila + 1, 3).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colFechaEsperadaETIA)), mfg.TextMatrix(Fila, colFechaEsperadaETIA))
        Hoja.Cells(Fila + 1, 4).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colDMAFinalPermit)), mfg.TextMatrix(Fila, colDMAFinalPermit))
        Hoja.Cells(Fila + 1, 5).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colEstado)), mfg.TextMatrix(Fila, colEstado))
        Hoja.Cells(Fila + 1, 6).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colYacimiento)), mfg.TextMatrix(Fila, colYacimiento))
        Hoja.Cells(Fila + 1, 7).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colPozo)), mfg.TextMatrix(Fila, colPozo))
        Hoja.Cells(Fila + 1, 8).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colFieldManifold)), mfg.TextMatrix(Fila, colFieldManifold))
        Hoja.Cells(Fila + 1, 9).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colBatteryAssigned)), mfg.TextMatrix(Fila, colBatteryAssigned))
        
      Else
        Hoja.Cells(Fila + 1, 1).Value = IIf(Fila = 0, UCase(mfg.TextMatrix(Fila, colWellID)), mfg.TextMatrix(Fila, colWellID))
      End If
      If Trim(mfg.TextMatrix(Fila, colDMAFinalPermit)) <> "" Then
        objExcel.Range("E" & Fila + 1).Select
        With objExcel.Selection.Interior
            .ColorIndex = 43
            .Pattern = xlSolid
        End With
      End If
      If Not IsNumeric(mfg.TextMatrix(Fila, colIDPozo)) And Fila > 0 Then
        With objExcel
            .Range("A" & Fila + 1 & ":" & "I" & Fila + 1).Select
            .Selection.Font.Bold = True
            .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            With .Selection.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            With .Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .ColorIndex = xlAutomatic
            End With
            .Selection.Borders(xlInsideVertical).LineStyle = xlNone
            With .Selection.Interior
                .ColorIndex = 15
                .Pattern = xlSolid
            End With
        End With
      End If
    Next Fila
    Form.Barra.Panels("info") = "Aplicando Formato..."
    AplicarMacroEstadoLineas objExcel, Hoja
    Libro.SaveAs (Form.cmd.FileName)
    Libro.Saved = True
    objExcel.Quit

ErrorHandler:
  If Err.Number <> 0 Then
  'Cierra la hoja y el la aplicación objExcel
    If Not Libro Is Nothing Then Libro.Saved = True
    If Not objExcel Is Nothing Then: objExcel.Quit
    'Liberar los objetos
    Set objExcel = Nothing
    Set Libro = Nothing
    Set Hoja = Nothing
    Err.Clear
  End If
End Sub


Private Sub AplicarMacroEstadoLineas(objExcel As Excel.Application, Hoja As Excel.Worksheet)
'Aplica el macro al estado de lineas

  With objExcel
    .Columns("A:I").Select
    .Columns("A:I").EntireColumn.AutoFit
    With .Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    .Range("A1:I1").Select
    With .Selection.Interior
        .ColorIndex = 10
        .Pattern = xlSolid
    End With
    .Selection.Font.ColorIndex = 2
    With .Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    .Selection.Font.Bold = True
    .Rows("1:1").RowHeight = 24
    .Application.WindowState = xlMinimized
    .Selection.Interior.ColorIndex = 51
    With .Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

  End With
End Sub



Private Sub ExportarTTMDatabase(Form As Form, mfg As MSHFlexGrid)
'Hace el exportar a la TTM database (Lo que antes hacia el boton del excel)
  On Error GoTo ErrorHandler

  Dim objSession As Object
  Dim objDatabase As Object
  
  Form.Barra.Panels("info") = "Conectando con la Base de datos ORACLE..."
  Set objSession = CreateObject("OracleInProcServer.XOraSession")
  Set objDatabase = objSession.OpenDatabase("AHPSPP", "TTM/TTM", 0)
  Dim Sql As String
  Dim i As Long
  
  Form.Barra.Panels("info") = "Generando secuencia de comandos SQL..."
  For i = 1 To mfg.Rows - 1
    If mfg.TextMatrix(i, colPozo) <> "" And mfg.TextMatrix(i, colFechaEntregaEIAxConsultoraAOXY) <> "" Then   ' Well, not Prospect
      Sql = Sql & "select '" & mfg.TextMatrix(i, colWellID) & "' well_name, to_date('" & mfg.TextMatrix(i, colFechaEntregaEIAxConsultoraAOXY) & "', 'DD/MM/YYYY HH24:MI') PREPARE_EIS, " & _
      "to_date('" & mfg.TextMatrix(i, colFechaPresentacionSMA) & "', 'DD/MM/YYYY HH24:MI') SUBMIT_EIS, to_date('" & mfg.TextMatrix(i, colDMAFinalPermit) & "', 'DD/MM/YYYY HH24:MI') ONSITE_DMA_INSPECTION, " & _
      "to_date('" & mfg.TextMatrix(i, colDMAFinalPermit) & "', 'DD/MM/YYYY HH24:MI') RECEIVE_DMA_APPROVAL, to_date('" & mfg.TextMatrix(i, colLandOwnerPermitDate) & "', 'DD/MM/YYYY HH24:MI') LAND_OWNER_PERMIT from dual union" & _
      vbCrLf
    End If
  Next
  ' Now strip off the final UNION stmt.
  Sql = Mid(Sql, 1, Len(Sql) - 6)
  ' Add the create or replace...
  Sql = "CREATE OR REPLACE VIEW TTM_LOAD_V AS " & Sql
  ' Create the Load View
  Form.Barra.Panels("info") = "Generando Vista..."
  objDatabase.ExecuteSQL (Sql)
  objDatabase.CommitTrans
  
    
  'Cross Tab the view:
  Form.Barra.Panels("info") = "Cruzando Vistas..."
  Dim sqlCT As String
  sqlCT = "CREATE OR REPLACE VIEW TTM_LOAD_CT_V AS SELECT well_name, decode(r, 1, 5, 2, 6, 3, 7, 4, 8, 5, 9) act_id, " & _
          "decode(r, 1, PREPARE_EIS, 2, SUBMIT_EIS, 3, ONSITE_DMA_INSPECTION, 4, RECEIVE_DMA_APPROVAL, 5, LAND_OWNER_PERMIT) val " & _
          "from TTM_LOAD_V x, (SELECT ROWNUM r FROM all_tables WHERE ROWNUM <= 5) b "
  
  ' Create the Cross tab View
  objDatabase.ExecuteSQL (sqlCT)
  objDatabase.CommitTrans
  
  '***************************************************************
  ' Merge into the Oracle database using the 2 views created above.
  Form.Barra.Panels("info") = "Combinando Tablas...."
  Dim sqlMerge As String
  sqlMerge = "MERGE INTO ttm_data d " & _
  "USING " & _
  "(SELECT * from TTM_LOAD_CT_V) s " & _
  "ON (d.well_name = s.well_name and d.activity_id = s.act_id) " & _
  "WHEN matched THEN " & _
  "        UPDATE SET d.value_dt = s.val " & _
  "WHEN NOT matched THEN " & _
  "     INSERT (ASSET_ID, WELL_NAME, ACTIVITY_ID, VALUE_DT) " & _
  "     VALUES (23, s.well_name, s.act_id, s.val) "
  
  ' Execute the Merge
  objDatabase.ExecuteSQL (sqlMerge)
  objDatabase.CommitTrans

ErrorHandler:
  ErrorHandler
End Sub

Private Sub ExportarPermittingData(Form As Form, mfg As MSHFlexGrid)
'Hace lo que hacia el boton Export Permitting Data del excel
  On Error GoTo ErrorHandler
  
  Form.Enabled = False
  Dim oCon As New ADODB.Connection
  oCon.ConnectionString = "provider=oraoledb.oracle.1;data source=AHPSPP;user id=TTM;password=TTM;"
  Form.Barra.Panels("info") = "Conectando..."
  oCon.Open
  Dim oCmd As New ADODB.Command
  With oCmd
      .ActiveConnection = oCon
      .CommandText = "TTM.TRUNC_TTM_PERMIT_DATA"
      .CommandType = adCmdStoredProc
      .Execute
  End With
  Set oCmd = Nothing
  oCon.Close
  Set oCon = Nothing
  InsertPermitData Form, mfg
  Form.Barra.Panels("info") = ""
  
ErrorHandler:
  ErrorHandler
End Sub


Private Sub InsertPermitData(Form As Form, mfg As MSHFlexGrid)
'Hace lo mismo que hacia la macro de excel
  On Error GoTo ErrorHandler

    Dim i As Long
    Dim sSql As String
    Dim iCounter As Long
    Dim oCon As New ADODB.Connection
    oCon.ConnectionString = "provider=oraoledb.oracle.1;data source=AHPSPP;user id=TTM;password=TTM;"
    oCon.Open
   
    
    sSql = "declare ary TTM_InsertPermittingData_pkg.permits_tbl; begin "
    iCounter = 1
    If mfg.Rows > 1 Then
      For i = 1 To mfg.Rows - 1
        Form.Barra.Panels("info") = "Procesando registro " & i & " de " & mfg.Rows - 1 & "..."
        If IsNumeric(mfg.TextMatrix(i, colIDPozo)) And mfg.TextMatrix(i, colWellID) <> "" And mfg.TextMatrix(i, colUbicacion) <> "GONE / DONE" Then
            sSql = sSql & "ary(" & iCounter & ").WELL_ID := '" & Replace(mfg.TextMatrix(i, colWellID), "'", "''") & "';"
            sSql = sSql & "ary(" & iCounter & ").WELL_PHASE := '" & Replace(mfg.TextMatrix(i, colUbicacion), "'", "''") & "';"
            sSql = sSql & "ary(" & iCounter & ").SITE_COMMENTS := '" & Replace(mfg.TextMatrix(i, colSiteVisit), "'", "''") & "';"
            If Not IsDate(mfg.TextMatrix(i, colFechaPrimerMonografia)) Then
              sSql = sSql & "ary(" & iCounter & ").MOBILIZE_SURVEYOR_DT := Null;"
              sSql = sSql & "ary(" & iCounter & ").SITE_STAKED_DT := Null;"
              sSql = sSql & "ary(" & iCounter & ").EIS_REQUESTED_DT := Null;"
            Else
              sSql = sSql & "ary(" & iCounter & ").MOBILIZE_SURVEYOR_DT := to_date('" & Format(DateAdd("d", -7, mfg.TextMatrix(i, colFechaPrimerMonografia)), "dd/mm/yyyy") & "','dd/MM/yyyy')" & ";"
              sSql = sSql & "ary(" & iCounter & ").SITE_STAKED_DT := to_date('" & Format(mfg.TextMatrix(i, colFechaPrimerMonografia), "dd/mm/yyyy") & "','dd/MM/yyyy')" & ";"
              sSql = sSql & "ary(" & iCounter & ").EIS_REQUESTED_DT := to_date('" & Format(mfg.TextMatrix(i, colFechaPrimerMonografia), "dd/mm/yyyy") & "','dd/MM/yyyy')" & ";"
            End If
            If Not IsDate(mfg.TextMatrix(i, colFechaUltimaMonografia)) Then
              sSql = sSql & "ary(" & iCounter & ").RESTAKED_DT := Null;"
            Else
              sSql = sSql & "ary(" & iCounter & ").RESTAKED_DT := to_date('" & Format(mfg.TextMatrix(i, colFechaUltimaMonografia), "dd/mm/yyyy") & "','dd/MM/yyyy')" & ";"
            End If
            If Not IsDate(mfg.TextMatrix(i, colDMAFinalPermit)) Then
              sSql = sSql & "ary(" & iCounter & ").DMA_APPROVED_DT := Null;"
            Else
              sSql = sSql & "ary(" & iCounter & ").DMA_APPROVED_DT := to_date('" & Format(mfg.TextMatrix(i, colDMAFinalPermit), "dd/mm/yyyy") & "','dd/MM/yyyy')" & ";"
            End If
            sSql = sSql & "ary(" & iCounter & ").STATUS := '" & mfg.TextMatrix(i, colStatus) & "';"
            If Not IsDate(mfg.TextMatrix(i, colLandOwnerPermitDate)) Then
              sSql = sSql & "ary(" & iCounter & ").LANDOWNER_PERMIT_DT := Null;"
            Else
              sSql = sSql & "ary(" & iCounter & ").LANDOWNER_PERMIT_DT := to_date('" & Format(mfg.TextMatrix(i, colLandOwnerPermitDate), "dd/mm/yyyy") & "','dd/MM/yyyy')" & ";"
            End If
            sSql = sSql & "ary(" & iCounter & ").FIELD_ABBR := '" & Replace(mfg.TextMatrix(i, colYacimiento), "'", "''") & "';"
            If Not IsDate(mfg.TextMatrix(i, colFechaEntregaEIAxConsultoraAOXY)) Then
              sSql = sSql & "ary(" & iCounter & ").EIS_PREPARED_DT := Null;"
            Else
              sSql = sSql & "ary(" & iCounter & ").EIS_PREPARED_DT := to_date('" & Format(mfg.TextMatrix(i, colFechaEntregaEIAxConsultoraAOXY), "dd/mm/yyyy") & "','dd/MM/yyyy')" & ";"
            End If
            If Not IsDate(mfg.TextMatrix(i, colFechaPresentacionSMA)) Then
              sSql = sSql & "ary(" & iCounter & ").EIS_SUBMITTED_DT := Null;"
            Else
              sSql = sSql & "ary(" & iCounter & ").EIS_SUBMITTED_DT := to_date('" & Format(mfg.TextMatrix(i, colFechaPresentacionSMA), "dd/mm/yyyy") & "','dd/MM/yyyy')" & ";"
            End If
            iCounter = iCounter + 1
            DoEvents
            If iCounter = 150 Then
              Form.Barra.Panels("info") = "Enviando grupo actual de pozos..."
              sSql = sSql & " TTM_InsertPermittingData_pkg.add_permits(ary); end;"
              oCon.Execute sSql
              sSql = "declare ary TTM_InsertPermittingData_pkg.permits_tbl; begin "
              iCounter = 1
            End If
        End If
      Next i
      Form.Barra.Panels("info") = "Enviando grupo actual de pozos..."
      sSql = sSql & " TTM_InsertPermittingData_pkg.add_permits(ary); end;"
      oCon.Execute sSql
      oCon.Execute "commit"
    End If
    oCon.Close
    Set oCon = Nothing
    
ErrorHandler:
  ErrorHandler
End Sub


