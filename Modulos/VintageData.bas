Attribute VB_Name = "VintageData"
Option Explicit
Public BDVintage As New ADODB.Connection
Public BDOXY As New ADODB.Connection
Public CargandoSolapa As Boolean


Public Sub ConectarBDVintageData()
'Realiza la conexion con el Vintage Data
  On Error GoTo ErrorHandler
  
  BDVintage.ConnectionString = "Provider = SQLOLEDB.1; Data Source=NAHWCSQL1\AH_SQL1; Initial Catalog=dataOxy; Integrated Security=SSPI;"
  BDVintage.Open
  
  
  BDOXY.Provider = "MSDAORA"
  BDOXY.ConnectionString = "Data Source=AHPSPP; User ID=oxy_read; Password=oxy_read; Persist Security Info=True;"
  BDOXY.Open
  
  
  
  
ErrorHandler:
  ErrorHandler
End Sub


'   13/01/2012
'   NO SE USA MAS, lee directo de DSinfo
'

'Public Sub ModificarPozosVintageData()
''Modifica los pozos del vintage data buscandolos y modificandolos si coinciden
'  On Error GoTo ErrorHandler
'  Dim RSVintage As New Recordset
'
'  Dim RSLocal As New Recordset
'  Dim ListaPozosModificados As String
'  Dim ListaPozosNoModificados As String
'  Dim Encontrado As Boolean
'  BD.BeginTrans
'
'  RSVintage.CursorLocation = adUseClient
'  RSVintage.Open "SELECT [Well ID], Rig, [Rig Order], Area, TD, [Total Days], [Remaining Days], [Start Date], Status, LandOwner, [End Date], EIADate, [Rig Order Checked] FROM IN_drilling_rigs_vw ORDER BY [Well ID]", BDVintage, adOpenStatic, adLockReadOnly, adCmdText
'  'trae todos los pozos del rig schedule de oxy data
'  While Not RSVintage.EOF
'
'    frmPozosFuturos.Barra.Panels("info") = "Analizando Well ID: " & RSVintage![WELL ID]
'
'    'busca cada pozo en este sistema , si existe le hace un update, sino lo agrega al listado de pozos nuevos
'    RSLocal.Open "SELECT * FROM POZOS WHERE UCASE(WELLID) = '" & UCase(RSVintage![WELL ID]) & "'", BD, adOpenDynamic, adLockOptimistic, adCmdText
'    If Not RSLocal.EOF Then
'      ListaPozosModificados = ListaPozosModificados & vbCrLf & RSVintage![WELL ID]
'      RSLocal!Equipo = RSVintage!Rig
'      RSLocal!Ubicacion = "RIGS SCHED."
'      RSLocal!OrdenUbicacion = ObtenerOrdenUbicacion("RIGS SCHED.")
'      RSLocal!RIGORDER = RSVintage![Rig Order]
'      RSLocal!Yacimiento = RSVintage!Area
'      RSLocal!TD = RSVintage!TD
'      RSLocal!TOTDAYS = RSVintage![Total Days]
'      RSLocal!REMDAYS = RSVintage![Remaining Days]
'      RSLocal!STARTDATE = RSVintage![Start Date]
'      RSLocal!Status = RSVintage!Status
'      RSLocal!LANDOWNER = RSVintage!LANDOWNER
'      RSLocal!ENDDATE = RSVintage![End Date]
'
'      RSLocal!FECHAESPERADAETIA = RSVintage!EIADate
'
''   If RSVintage![WELL ID] = "EH-3177" Then
''   MsgBox "ok"
''   End If
'
'      RSLocal!RIGORDERCHECKED = CBool(RSVintage![Rig Order Checked])
'
'      RSLocal.Update
'      'Grabo historial
'
'      'vergpTEST
''      MsgBox "13"
'      BD.Execute "INSERT INTO HIST_POZOS (WELLID, Ubicacion, Equipo, MONOGRAFIA, ACUIFERO, Prognosis, Programa, Yacimiento, TIPOYACIMIENTO, Pozo, Prospect, FECHAPRIMERMONOGRAFIA, MONOGRAFIAS, FECHAULTIMAMONOGRAFIA, WELLINFORMED, INFORMEDBY, Definitiva, XPDC, YPDC, XPOS94, YPOS94, " & _
'             " FECHASOLICITUD, FECHAPRIORIDAD, DOCUMENTOAPREPARAR, FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO, FIRSTPROD, " & _
'             " TD, TOTDAYS, REMDAYS, Status, STARTDATE, ENDDATE, LANDOWNER, LANDOWNERPERMITDATE, " & _
'             " Consult, Type, FechaPedidoEtia, FECHAESPERADAETIA, " & _
'             " CONSULTANTRECOMENDATION, DMAPERMIT, Estado, " & _
'             " IDMANIFIESTO, FECHAMANIFIESTO, FECHAENTREGAEIAXCONSULTORAOXY, EIAPRESENTADO, FECHAENVIOACS, FECHAPRESENTACIONDMA, FECHAPRESENTACIONSMA, PAGOTASAADMINISTRATIVA, FECHAINFOCOMPLEMENTARIA, TECHNICALREPORT, TASAADMINISTRATIVA, TASACONTRALOR, Estudio, Adenda, FECHAINICIODIA, FECHAFINDIA, INFORMEDEAVANCEDEOBRA50PORCIENTO, INFORMEDEAVANCEDEOBRA100PORCIENTO, INFORMEEVALUACIONARQUEOLOGICO, " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA, TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA, TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA, TIEMPOENTREPRESENTACIONANTEDMAYVISITA, TIEMPOENTREVISITAYAPROBACIONFINALDEDMA, TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA, TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA, TIEMPOENTREENTREGAEIADMAYAPROBACION, TipoMovimiento, XID, XFECHA, XUSUARIO, RIGORDERCHECKED)" & _
'             " SELECT WELLID AS [Well ID], Ubicacion, Equipo, MONOGRAFIA AS [Monografía], ACUIFERO AS [Acuífero], Prognosis AS Prognosis, Programa AS Programa, Yacimiento, TIPOYACIMIENTO AS [Tipo Yacimiento], Pozo, Prospect, FECHAPRIMERMONOGRAFIA AS [Fecha Primera Monografia], MONOGRAFIAS AS [Monografías], FECHAULTIMAMONOGRAFIA AS [Fecha Ultima Monografía],  WELLINFORMED AS [Well Informed], INFORMEDBY AS [Informed By], Definitiva AS Definitiva, XPDC AS [X PDC], YPDC AS [Y PDC], XPOS94 AS [X POS94], YPOS94 AS [Y POS94], " & _
'             " FECHASOLICITUD AS [Fecha Solicitud], FECHAPRIORIDAD as [Fecha Prioridad], DOCUMENTOAPREPARAR AS [Documento a Preparar], FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO AS [Fecha Entrega Dictamen Tco], FIRSTPROD AS [First Prod], " & _
'             " TD, TOTDAYS AS [Tot Days], REMDAYS AS [Rem Days], Status, STARTDATE AS [Start Date], ENDDATE AS [End Date], LANDOWNER AS [Land Owner], LANDOWNERPERMITDATE AS [Land Owner Permit Date], " & _
'             " Consult, Type, FechaPedidoEtia AS [Fecha Pedido ETIA], FechaEsperadaEtia AS [Fecha Esperada ETIA], " & _
'             " CONSULTANTRECOMENDATION AS [Consultant Recomendation], DMAPERMIT AS [DMA Permit], Estado, " & _
'             " IDMANIFIESTO AS [ID Manifiesto], FECHAMANIFIESTO AS [Fecha Manifiesto], FECHAENTREGAEIAXCONSULTORAOXY AS [Fecha Entrega EIA Por Consultor a OXY], EIAPRESENTADO AS [EIA Presentado], FECHAENVIOACS AS [Fecha Envio ACS], FECHAPRESENTACIONDMA AS [Fecha Presentación DMA], FECHAPRESENTACIONSMA AS [Fecha Presentación SMA], PAGOTASAADMINISTRATIVA AS [Pago Tasa Administrativa], FECHAINFOCOMPLEMENTARIA AS [Fecha Info Complementaria], TECHNICALREPORT AS [Technical Report], TASAADMINISTRATIVA AS [Tasa Administrativa], TASACONTRALOR AS [Tasa Contralor], Estudio, Adenda, FECHAINICIODIA AS [Fecha Inicio DIA], FECHAFINDIA AS [Fecha Fin DIA], INFORMEDEAVANCEDEOBRA50PORCIENTO AS [Informe de Avance de Obra 50%], INFORMEDEAVANCEDEOBRA100PORCIENTO AS [Informe de Avance de Obra 100%], INFORMEEVALUACIONARQUEOLOGICO AS [Informe Evaluacion Arqueologico], " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA AS [Tiempo Entre Pedido ETIA y Recepcion ETIA], TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA AS [Tiempo Entre Primer Monografía y Recepción ETIA], TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA AS [Tiempo Entre Recepcion ETIA y Presentación ante DMA], TIEMPOENTREPRESENTACIONANTEDMAYVISITA AS [Tiempo Entre Presentación Ante DMA y Visita], TIEMPOENTREVISITAYAPROBACIONFINALDEDMA AS [Tiempo Entre Visita y Aprobación Final De DMA], TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA AS [Tiempo Entre Presentación De Pozo y Aprobación Final DMA], TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA AS [Tiempo Entre Primer Monografia y Pedido EIA], TIEMPOENTREENTREGAEIADMAYAPROBACION AS [Tiempo Entre Entrega EIA a DMA y Aprobacion], 'M', IDPOZO, #" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "#,'Vintage Data', RIGORDERCHECKED FROM POZOS WHERE IDPOZO IN (" & RSLocal!IDPozo & ")"
'    Else
'      ListaPozosNoModificados = ListaPozosNoModificados & vbCrLf & RSVintage![WELL ID]
''      If RSVintage![WELL ID] = "EH-3177" Then
''        MsgBox "ok"
''      End If
'    End If
'    RSLocal.Close
'    RSVintage.MoveNext
'  Wend
'  RSVintage.Close
'
'  'ahora actualiza laas cordenadas der los pozos
'
'  RSVintage.Open "SELECT [WellName], Area, X, Y, TD, LandOwner FROM IN_drilling_Inventory_vw where rig = '(ninguno)' ORDER BY WELLNAME", BDVintage, adOpenStatic, adLockReadOnly, adCmdText
'  'RSVintage.Open "SELECT [WellName], Area, X, Y, TD, LandOwner, Status FROM IN_drilling_Inventory_vw where rig = '(ninguno)' ORDER BY WELLNAME", BDVintage, adOpenStatic, adLockReadOnly, adCmdText
'  While Not RSVintage.EOF
'    frmPozosFuturos.Barra.Panels("info") = "Analizando Well ID: " & RSVintage![WELLNAME]
'    RSLocal.Open "SELECT * FROM POZOS WHERE UCASE(WELLID) = '" & UCase(RSVintage![WELLNAME]) & "'", BD, adOpenDynamic, adLockOptimistic, adCmdText
'    If Not RSLocal.EOF Then
'      ListaPozosModificados = ListaPozosModificados & vbCrLf & RSVintage![WELLNAME]
'      RSLocal!Yacimiento = RSVintage!Area
'      RSLocal!XPOS94 = RSVintage!X
'      RSLocal!YPOS94 = RSVintage!Y
'      RSLocal!TD = RSVintage!TD
'     ' RSLocal!Status = RSVintage!Status
'      RSLocal!LANDOWNER = RSVintage!LANDOWNER
'      RSLocal!Ubicacion = "DRILLING INV."
'      RSLocal!OrdenUbicacion = ObtenerOrdenUbicacion("DRILLING INV.")
'
'      RSLocal!RIGORDER = 999
'      RSLocal!RIGORDERCHECKED = False
'
'      'actualizo el status
'      Dim RSVintage2 As New Recordset
'      RSVintage2.Open "SELECT [Status] FROM ubiFac2_vw where [Well ID] = '" & RSVintage![WELLNAME] & "'", BDVintage, adOpenStatic, adLockReadOnly, adCmdText
'     ' RSVintage2.CursorLocation = adUseClient
'      If Not RSVintage2.EOF Then
'        RSLocal!Status = RSVintage2!Status
'      End If
'      RSVintage2.Close
'
'      RSLocal.Update
'      'Grabo Historial
'
'      'vergpTEST
'
''MsgBox "14"
'      BD.Execute "INSERT INTO HIST_POZOS (WELLID, Ubicacion, Equipo, MONOGRAFIA, ACUIFERO, Prognosis, Programa, Yacimiento, TIPOYACIMIENTO, Pozo, Prospect, FECHAPRIMERMONOGRAFIA, MONOGRAFIAS, FECHAULTIMAMONOGRAFIA, WELLINFORMED, INFORMEDBY, Definitiva, XPDC, YPDC, XPOS94, YPOS94, " & _
'             " FECHASOLICITUD, FECHAPRIORIDAD, DOCUMENTOAPREPARAR, FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO, FIRSTPROD, " & _
'             " TD, TOTDAYS, REMDAYS, Status, STARTDATE, ENDDATE, LANDOWNER, LANDOWNERPERMITDATE, " & _
'             " Consult, Type, FechaPedidoEtia, FECHAESPERADAETIA, " & _
'             " CONSULTANTRECOMENDATION, DMAPERMIT, Estado, " & _
'             " IDMANIFIESTO, FECHAMANIFIESTO, FECHAENTREGAEIAXCONSULTORAOXY, EIAPRESENTADO, FECHAENVIOACS, FECHAPRESENTACIONDMA, FECHAPRESENTACIONSMA, PAGOTASAADMINISTRATIVA, FECHAINFOCOMPLEMENTARIA, TECHNICALREPORT, TASAADMINISTRATIVA, TASACONTRALOR, Estudio, Adenda, FECHAINICIODIA, FECHAFINDIA, INFORMEDEAVANCEDEOBRA50PORCIENTO, INFORMEDEAVANCEDEOBRA100PORCIENTO, INFORMEEVALUACIONARQUEOLOGICO, " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA, TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA, TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA, TIEMPOENTREPRESENTACIONANTEDMAYVISITA, TIEMPOENTREVISITAYAPROBACIONFINALDEDMA, TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA, TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA, TIEMPOENTREENTREGAEIADMAYAPROBACION, TipoMovimiento, XID, XFECHA, XUSUARIO, RIGORDERCHECKED)" & _
'             " SELECT WELLID AS [Well ID], Ubicacion, Equipo, MONOGRAFIA AS [Monografía], ACUIFERO AS [Acuífero], Prognosis AS Prognosis, Programa AS Programa, Yacimiento, TIPOYACIMIENTO AS [Tipo Yacimiento], Pozo, Prospect, FECHAPRIMERMONOGRAFIA AS [Fecha Primera Monografia], MONOGRAFIAS AS [Monografías], FECHAULTIMAMONOGRAFIA AS [Fecha Ultima Monografía],  WELLINFORMED AS [Well Informed], INFORMEDBY AS [Informed By], Definitiva AS Definitiva, XPDC AS [X PDC], YPDC AS [Y PDC], XPOS94 AS [X POS94], YPOS94 AS [Y POS94], " & _
'             " FECHASOLICITUD AS [Fecha Solicitud], FECHAPRIORIDAD as [Fecha Prioridad], DOCUMENTOAPREPARAR AS [Documento a Preparar], FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO AS [Fecha Entrega Dictamen Tco], FIRSTPROD AS [First Prod], " & _
'             " TD, TOTDAYS AS [Tot Days], REMDAYS AS [Rem Days], Status, STARTDATE AS [Start Date], ENDDATE AS [End Date], LANDOWNER AS [Land Owner], LANDOWNERPERMITDATE AS [Land Owner Permit Date], " & _
'             " Consult, Type, FechaPedidoEtia AS [Fecha Pedido ETIA], FechaEsperadaEtia AS [Fecha Esperada ETIA], " & _
'             " CONSULTANTRECOMENDATION AS [Consultant Recomendation], DMAPERMIT AS [DMA Permit], Estado, " & _
'             " IDMANIFIESTO AS [ID Manifiesto], FECHAMANIFIESTO AS [Fecha Manifiesto], FECHAENTREGAEIAXCONSULTORAOXY AS [Fecha Entrega EIA Por Consultor a OXY], EIAPRESENTADO AS [EIA Presentado], FECHAENVIOACS AS [Fecha Envio ACS], FECHAPRESENTACIONDMA AS [Fecha Presentación DMA], FECHAPRESENTACIONSMA AS [Fecha Presentación SMA], PAGOTASAADMINISTRATIVA AS [Pago Tasa Administrativa], FECHAINFOCOMPLEMENTARIA AS [Fecha Info Complementaria], TECHNICALREPORT AS [Technical Report], TASAADMINISTRATIVA AS [Tasa Administrativa], TASACONTRALOR AS [Tasa Contralor], Estudio, Adenda, FECHAINICIODIA AS [Fecha Inicio DIA], FECHAFINDIA AS [Fecha Fin DIA], INFORMEDEAVANCEDEOBRA50PORCIENTO AS [Informe de Avance de Obra 50%], INFORMEDEAVANCEDEOBRA100PORCIENTO AS [Informe de Avance de Obra 100%], INFORMEEVALUACIONARQUEOLOGICO AS [Informe Evaluacion Arqueologico], " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA AS [Tiempo Entre Pedido ETIA y Recepcion ETIA], TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA AS [Tiempo Entre Primer Monografía y Recepción ETIA], TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA AS [Tiempo Entre Recepcion ETIA y Presentación ante DMA], TIEMPOENTREPRESENTACIONANTEDMAYVISITA AS [Tiempo Entre Presentación Ante DMA y Visita], TIEMPOENTREVISITAYAPROBACIONFINALDEDMA AS [Tiempo Entre Visita y Aprobación Final De DMA], TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA AS [Tiempo Entre Presentación De Pozo y Aprobación Final DMA], TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA AS [Tiempo Entre Primer Monografia y Pedido EIA], TIEMPOENTREENTREGAEIADMAYAPROBACION AS [Tiempo Entre Entrega EIA a DMA y Aprobacion], 'M', IDPOZO, #" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "#,'Vintage Data', RIGORDERCHECKED FROM POZOS WHERE IDPOZO IN (" & RSLocal!IDPozo & ")"
'    Else
'      ListaPozosNoModificados = ListaPozosNoModificados & vbCrLf & RSVintage![WELLNAME]
'    End If
'    RSLocal.Close
'    RSVintage.MoveNext
'  Wend
'  RSVintage.Close
'  BD.CommitTrans
'  BD.BeginTrans
'  frmPozosFuturos.Barra.Panels("info") = "Enviando pozos a Gone/Done"
'  RSVintage.CursorLocation = adUseClient
'
'  'ahora recorre mis pozos y  verfica q continuen dentro de rig sched. caso contrario lo manda a gone done
'  'q pasa si el pozo no estaba entre los primeros 7 y decidieron sacarlo?, ahora se manda a gone done
'
'
'  RSLocal.Open "SELECT * FROM POZOS WHERE UBICACION = 'RIGS SCHED.' ORDER BY WELLID", BD, adOpenDynamic, adLockOptimistic, adCmdText
'  RSVintage.Open "SELECT [Well ID] FROM IN_drilling_rigs_vw ORDER BY [Well ID]", BDVintage, adOpenDynamic, adLockReadOnly, adCmdText
'  While Not RSLocal.EOF
'    Encontrado = False
'    RSVintage.MoveFirst
'    While Not RSVintage.EOF And Not Encontrado
'      If UCase(RSVintage![WELL ID]) <> UCase(RSLocal!WellID) Then
'        RSVintage.MoveNext
'      Else
'        Encontrado = True
'      End If
'    Wend
'    If RSVintage.EOF Then
'      RSLocal!Ubicacion = "GONE / DONE"
'      RSLocal!OrdenUbicacion = ObtenerOrdenUbicacion("GONE / DONE")
'      RSLocal!RIGORDER = -1
'      RSLocal!RIGORDERCHECKED = False
'      RSLocal.Update
'      'Grabo historial
'
'      'vergpTEST
''      MsgBox "15"
'
'      BD.Execute "INSERT INTO HIST_POZOS (WELLID, Ubicacion, Equipo, MONOGRAFIA, ACUIFERO, Prognosis, Programa, Yacimiento, TIPOYACIMIENTO, Pozo, Prospect, FECHAPRIMERMONOGRAFIA, MONOGRAFIAS, FECHAULTIMAMONOGRAFIA, WELLINFORMED, INFORMEDBY, Definitiva, XPDC, YPDC, XPOS94, YPOS94, " & _
'             " FECHASOLICITUD, FECHAPRIORIDAD, DOCUMENTOAPREPARAR, FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO, FIRSTPROD, " & _
'             " TD, TOTDAYS, REMDAYS, Status, STARTDATE, ENDDATE, LANDOWNER, LANDOWNERPERMITDATE, " & _
'             " Consult, Type, FechaPedidoEtia, FECHAESPERADAETIA, " & _
'             " CONSULTANTRECOMENDATION, DMAPERMIT, Estado, " & _
'             " IDMANIFIESTO, FECHAMANIFIESTO, FECHAENTREGAEIAXCONSULTORAOXY, EIAPRESENTADO, FECHAENVIOACS, FECHAPRESENTACIONDMA, FECHAPRESENTACIONSMA, PAGOTASAADMINISTRATIVA, FECHAINFOCOMPLEMENTARIA, TECHNICALREPORT, TASAADMINISTRATIVA, TASACONTRALOR, Estudio, Adenda, FECHAINICIODIA, FECHAFINDIA, INFORMEDEAVANCEDEOBRA50PORCIENTO, INFORMEDEAVANCEDEOBRA100PORCIENTO, INFORMEEVALUACIONARQUEOLOGICO, " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA, TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA, TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA, TIEMPOENTREPRESENTACIONANTEDMAYVISITA, TIEMPOENTREVISITAYAPROBACIONFINALDEDMA, TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA, TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA, TIEMPOENTREENTREGAEIADMAYAPROBACION, TipoMovimiento, XID, XFECHA, XUSUARIO, RIGORDERCHECKED)" & _
'             " SELECT WELLID AS [Well ID], Ubicacion, Equipo, MONOGRAFIA AS [Monografía], ACUIFERO AS [Acuífero], Prognosis AS Prognosis, Programa AS Programa, Yacimiento, TIPOYACIMIENTO AS [Tipo Yacimiento], Pozo, Prospect, FECHAPRIMERMONOGRAFIA AS [Fecha Primera Monografia], MONOGRAFIAS AS [Monografías], FECHAULTIMAMONOGRAFIA AS [Fecha Ultima Monografía],  WELLINFORMED AS [Well Informed], INFORMEDBY AS [Informed By], Definitiva AS Definitiva, XPDC AS [X PDC], YPDC AS [Y PDC], XPOS94 AS [X POS94], YPOS94 AS [Y POS94], " & _
'             " FECHASOLICITUD AS [Fecha Solicitud], FECHAPRIORIDAD as [Fecha Prioridad], DOCUMENTOAPREPARAR AS [Documento a Preparar],FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO AS [Fecha Entrega Dictamen Tco], FIRSTPROD AS [First Prod], " & _
'             " TD, TOTDAYS AS [Tot Days], REMDAYS AS [Rem Days], Status, STARTDATE AS [Start Date], ENDDATE AS [End Date], LANDOWNER AS [Land Owner], LANDOWNERPERMITDATE AS [Land Owner Permit Date], " & _
'             " Consult, Type, FechaPedidoEtia AS [Fecha Pedido ETIA], FechaEsperadaEtia AS [Fecha Esperada ETIA], " & _
'             " CONSULTANTRECOMENDATION AS [Consultant Recomendation], DMAPERMIT AS [DMA Permit], Estado, " & _
'             " IDMANIFIESTO AS [ID Manifiesto], FECHAMANIFIESTO AS [Fecha Manifiesto], FECHAENTREGAEIAXCONSULTORAOXY AS [Fecha Entrega EIA Por Consultor a OXY], EIAPRESENTADO AS [EIA Presentado], FECHAENVIOACS AS [Fecha Envio ACS], FECHAPRESENTACIONDMA AS [Fecha Presentación DMA], FECHAPRESENTACIONSMA AS [Fecha Presentación SMA], PAGOTASAADMINISTRATIVA AS [Pago Tasa Administrativa], FECHAINFOCOMPLEMENTARIA AS [Fecha Info Complementaria], TECHNICALREPORT AS [Technical Report], TASAADMINISTRATIVA AS [Tasa Administrativa], TASACONTRALOR AS [Tasa Contralor], Estudio, Adenda, FECHAINICIODIA AS [Fecha Inicio DIA], FECHAFINDIA AS [Fecha Fin DIA], INFORMEDEAVANCEDEOBRA50PORCIENTO AS [Informe de Avance de Obra 50%], INFORMEDEAVANCEDEOBRA100PORCIENTO AS [Informe de Avance de Obra 100%], INFORMEEVALUACIONARQUEOLOGICO AS [Informe Evaluacion Arqueologico], " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA AS [Tiempo Entre Pedido ETIA y Recepcion ETIA], TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA AS [Tiempo Entre Primer Monografía y Recepción ETIA], TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA AS [Tiempo Entre Recepcion ETIA y Presentación ante DMA], TIEMPOENTREPRESENTACIONANTEDMAYVISITA AS [Tiempo Entre Presentación Ante DMA y Visita], TIEMPOENTREVISITAYAPROBACIONFINALDEDMA AS [Tiempo Entre Visita y Aprobación Final De DMA], TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA AS [Tiempo Entre Presentación De Pozo y Aprobación Final DMA], TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA AS [Tiempo Entre Primer Monografia y Pedido EIA], TIEMPOENTREENTREGAEIADMAYAPROBACION AS [Tiempo Entre Entrega EIA a DMA y Aprobacion], 'M', IDPOZO, #" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "#,'Vintage Data', RIGORDERCHECKED FROM POZOS WHERE IDPOZO IN (" & RSLocal!IDPozo & ")"
'    End If
'    RSLocal.MoveNext
'  Wend
'  frmPozosFuturos.Barra.Panels("info") = ""
'  ListaPozosModificados = Mid(ListaPozosModificados, 3)
'  ListaPozosNoModificados = Mid(ListaPozosNoModificados, 3)
'  frmResumenImportacion.AsignarEncontrados ListaPozosModificados
'  frmResumenImportacion.AsignarnoEncontrados ListaPozosNoModificados
'  frmPozosFuturos.Barra.Panels("info") = "Actualizando Fecha 'First Prod'"
'  CargarFechasFirstProd
'  frmResumenImportacion.Show
'
'
'
'ErrorHandler:
'  If Err.Number = 0 Then
'    BD.CommitTrans
'  Else
'    ErrorHandler
'    BD.RollbackTrans
'  End If
'End Sub



Private Sub CargarFechasFirstProd()
  Dim objSession As Object
  Dim objDatabase As Object
  Dim rs As New Recordset
  On Error GoTo ErrorHandler
  
  Dim RSLocal As New Recordset
  Dim RSOXY As New Recordset
  
'  RSLocal.Open "SELECT * FROM POZOS where Ubicacion = 'GONE / DONE'", BD, adOpenDynamic, adLockOptimistic, adCmdText
'  Do While Not RSLocal.EOF
'    frmPozosFuturos.Barra.Panels("info") = "Actualizando Fecha First Prod en Pozo: " & RSLocal![WellID]
'    Dim strFechaPrimerProd As String
'    strFechaPrimerProd = ""
'      'calculo fecha de primer produccion
'      'If UCase(RSLocal![WellID]) = "LH-2042" Then
'        RSOXY.Open "SELECT VALUE_DT From DASHBD.TTM_DATA WHERE ACTIVITY_ID = 14 and WELL_NAME = '" & UCase(RSLocal![WellID]) & "'", BDOXY, adOpenDynamic, adLockReadOnly, adCmdText
'        If Not RSOXY.EOF Then
'          strFechaPrimerProd = Format(RSOXY!VALUE_DT, "dd/mm/yyyy")
'           'fecha de primer produccion
'          If IsDate(strFechaPrimerProd) Then
'            RSLocal!FIRSTPROD = strFechaPrimerProd
'            RSLocal.Update
'          End If
'        End If
'        RSOXY.Close
'     ' End If
'      RSLocal.MoveNext
'  Loop
'  RSLocal.Close
  
  
  
''   13/01/2012
''   NO SE USA MAS, lee directo de DSinfo
''
'  RSOXY.Open "SELECT WELL_NAME,VALUE_DT From TTM.TTM_DATA WHERE ACTIVITY_ID = 27 and VALUE_DT is not null", BDOXY, adOpenDynamic, adLockReadOnly, adCmdText
'
'  Do While Not RSOXY.EOF
'    frmPozosFuturos.Barra.Panels("info") = "Actualizando Fecha First Prod en Pozo: " & RSOXY![WELL_NAME]
'    Dim strFechaPrimerProd As String
'    strFechaPrimerProd = ""
'    strFechaPrimerProd = Format(RSOXY!VALUE_DT, "dd/mm/yyyy")
'    If IsDate(strFechaPrimerProd) Then
'        RSLocal.Open "SELECT * FROM POZOS where WELLID = '" & UCase(RSOXY![WELL_NAME]) & "'", BD, adOpenDynamic, adLockOptimistic, adCmdText
'        If Not RSLocal.EOF Then
'            RSLocal!FIRSTPROD = strFechaPrimerProd
'            RSLocal.Update
'        End If
'        RSLocal.Close
'    End If
'    RSOXY.MoveNext
'  Loop
'  RSOXY.Close
    
    
    
'
'  frmPozosFuturos.Barra.Panels("info") = "Obteniendo Fechas FirstProd"
'  Set objSession = CreateObject("OracleInProcServer.XOraSession")
'  Set objDatabase = objSession.OpenDatabase("oxy", "dashbd/dashbd", 0)
'  RS.CursorLocation = adUseClient
'  RS.Open "SELECT * FROM ", objDatabase, adOpenStatic, adLockReadOnly, adCmdText
'
  
ErrorHandler:
  If Err.Number <> 0 Then Err.Raise Err.Number
End Sub


'   13/01/2012
'   NO SE USA MAS, lee directo de DSinfo
'
'
'Public Sub CrearPozosNoEncontrados(IDs As String)
''Cierra la conexion con la base de datos del vintage data
'  On Error GoTo ErrorHandler
'  Dim RSVintage As New Recordset
'  Dim RSLocal As New Recordset
'
'  BD.BeginTrans
'  RSVintage.CursorLocation = adUseClient
'  RSLocal.CursorLocation = adUseClient
'  'crea los pozos que no encontro
'  'puede haber pozos despues de los primeros 7 reg
'
'
'
'  RSVintage.Open "SELECT [Well ID], Rig, [Rig Order], Area, TD, [Total Days], [Remaining Days], [Start Date], Status, LandOwner, [End Date], EIADate, [Rig Order Checked] FROM IN_drilling_rigs_vw WHERE [Well ID] IN (" & IDs & ")", BDVintage, adOpenStatic, adLockReadOnly, adCmdText
'  While Not RSVintage.EOF
'    If Not ExistePozo(RSVintage![WELL ID]) Then
'      frmPozosFuturos.Barra.Panels("info") = "Generando pozo Well ID: " & RSVintage![WELL ID]
'      RSLocal.Open "POZOS", BD, adOpenStatic, adLockOptimistic, adCmdTable
'      RSLocal.AddNew
'      RSLocal!WellID = RSVintage![WELL ID]
'      RSLocal!Equipo = RSVintage!Rig
'      RSLocal!Ubicacion = "RIGS SCHED."
'      RSLocal!OrdenUbicacion = ObtenerOrdenUbicacion("RIGS SCHED.")
'      RSLocal!RIGORDER = RSVintage![Rig Order]
'      RSLocal!Yacimiento = RSVintage!Area
'      RSLocal!TD = RSVintage!TD
'      RSLocal!TOTDAYS = RSVintage![Total Days]
'      RSLocal!REMDAYS = RSVintage![Remaining Days]
'      RSLocal!STARTDATE = RSVintage![Start Date]
'      RSLocal!Status = RSVintage!Status
'      RSLocal!LANDOWNER = RSVintage!LANDOWNER
'      RSLocal!ENDDATE = RSVintage![End Date]
'      RSLocal!FECHAESPERADAETIA = RSVintage!EIADate
'      RSLocal!RIGORDERCHECKED = CBool(RSVintage![Rig Order Checked])
'      RSLocal.Update
'
'      'vergpTEST
''      MsgBox "16"
'      BD.Execute "INSERT INTO HIST_POZOS (WELLID, Ubicacion, Equipo, MONOGRAFIA, ACUIFERO, Prognosis, Programa, Yacimiento, TIPOYACIMIENTO, Pozo, Prospect,FECHAPRIMERMONOGRAFIA, MONOGRAFIAS, FECHAULTIMAMONOGRAFIA, WELLINFORMED, INFORMEDBY, Definitiva, XPDC, YPDC, XPOS94, YPOS94, " & _
'             " FECHASOLICITUD, FECHAPRIORIDAD, DOCUMENTOAPREPARAR, FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO, FIRSTPROD, " & _
'             " TD, TOTDAYS, REMDAYS, Status, STARTDATE, ENDDATE, LANDOWNER, LANDOWNERPERMITDATE, " & _
'             " Consult, Type, FechaPedidoEtia, FECHAESPERADAETIA, " & _
'             " CONSULTANTRECOMENDATION, DMAPERMIT, Estado, " & _
'             " IDMANIFIESTO, FECHAMANIFIESTO, FECHAENTREGAEIAXCONSULTORAOXY, EIAPRESENTADO, FECHAENVIOACS, FECHAPRESENTACIONDMA, FECHAPRESENTACIONSMA, PAGOTASAADMINISTRATIVA, FECHAINFOCOMPLEMENTARIA, TECHNICALREPORT, TASAADMINISTRATIVA, TASACONTRALOR, Estudio, Adenda, FECHAINICIODIA, FECHAFINDIA,INFORMEDEAVANCEDEOBRA50PORCIENTO, INFORMEDEAVANCEDEOBRA100PORCIENTO, INFORMEEVALUACIONARQUEOLOGICO, " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA, TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA, TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA, TIEMPOENTREPRESENTACIONANTEDMAYVISITA, TIEMPOENTREVISITAYAPROBACIONFINALDEDMA, TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA, TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA, TIEMPOENTREENTREGAEIADMAYAPROBACION, TipoMovimiento, XID, XFECHA, XUSUARIO, RIGORDERCHECKED)" & _
'             " SELECT WELLID AS [Well ID], Ubicacion, Equipo, MONOGRAFIA AS [Monografía], ACUIFERO AS [Acuífero], Prognosis AS Prognosis, Programa AS Programa, Yacimiento, TIPOYACIMIENTO AS [Tipo Yacimiento], Pozo, Prospect, FECHAPRIMERMONOGRAFIA AS [Fecha Primera Monografia], MONOGRAFIAS AS [Monografías], FECHAULTIMAMONOGRAFIA AS [Fecha Ultima Monografía],  WELLINFORMED AS [Well Informed], INFORMEDBY AS [Informed By], Definitiva AS Definitiva, XPDC AS [X PDC], YPDC AS [Y PDC], XPOS94 AS [X POS94], YPOS94 AS [Y POS94], " & _
'             " FECHASOLICITUD AS [Fecha Solicitud], FECHAPRIORIDAD as [Fecha Prioridad], DOCUMENTOAPREPARAR AS [Documento a Preparar],FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO AS [Fecha Entrega Dictamen Tco], FIRSTPROD AS [First Prod], " & _
'             " TD, TOTDAYS AS [Tot Days], REMDAYS AS [Rem Days], Status, STARTDATE AS [Start Date], ENDDATE AS [End Date], LANDOWNER AS [Land Owner], LANDOWNERPERMITDATE AS [Land Owner Permit Date], " & _
'             " Consult, Type, FechaPedidoEtia AS [Fecha Pedido ETIA], FechaEsperadaEtia AS [Fecha Esperada ETIA], " & _
'             " CONSULTANTRECOMENDATION AS [Consultant Recomendation], DMAPERMIT AS [DMA Permit], Estado, " & _
'             " IDMANIFIESTO AS [ID Manifiesto], FECHAMANIFIESTO AS [Fecha Manifiesto], FECHAENTREGAEIAXCONSULTORAOXY AS [Fecha Entrega EIA Por Consultor a OXY], EIAPRESENTADO AS [EIA Presentado], FECHAENVIOACS AS [Fecha Envio ACS], FECHAPRESENTACIONDMA AS [Fecha Presentación DMA], FECHAPRESENTACIONSMA AS [Fecha Presentación SMA], PAGOTASAADMINISTRATIVA AS [Pago Tasa Administrativa], FECHAINFOCOMPLEMENTARIA AS [Fecha Info Complementaria], TECHNICALREPORT AS [Technical Report], TASAADMINISTRATIVA AS [Tasa Administrativa], TASACONTRALOR AS [Tasa Contralor], Estudio, Adenda, FECHAINICIODIA AS [Fecha Inicio DIA], FECHAFINDIA AS [Fecha Fin DIA], INFORMEDEAVANCEDEOBRA50PORCIENTO AS [Informe de Avance de Obra 50%], INFORMEDEAVANCEDEOBRA100PORCIENTO AS [Informe de Avance de Obra 100%], INFORMEEVALUACIONARQUEOLOGICO AS [Informe Evaluacion Arqueologico], " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA AS [Tiempo Entre Pedido ETIA y Recepcion ETIA], TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA AS [Tiempo Entre Primer Monografía y Recepción ETIA], TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA AS [Tiempo Entre Recepcion ETIA y Presentación ante DMA], TIEMPOENTREPRESENTACIONANTEDMAYVISITA AS [Tiempo Entre Presentación Ante DMA y Visita], TIEMPOENTREVISITAYAPROBACIONFINALDEDMA AS [Tiempo Entre Visita y Aprobación Final De DMA], TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA AS [Tiempo Entre Presentación De Pozo y Aprobación Final DMA], TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA AS [Tiempo Entre Primer Monografia y Pedido EIA], TIEMPOENTREENTREGAEIADMAYAPROBACION AS [Tiempo Entre Entrega EIA a DMA y Aprobacion], 'A', IDPOZO, #" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "#,'Vintage Data', RIGORDERCHECKED FROM POZOS WHERE IDPOZO IN (" & RSLocal!IDPozo & ")"
'      RSLocal.Close
'
'    End If
'
'    RSVintage.MoveNext
'
'  Wend
'
'  RSVintage.Close
'
'  'actualiza las coordenadas
'  RSVintage.Open "SELECT [WellName], Area, X, Y, TD, LandOwner FROM IN_drilling_Inventory_vw WHERE rig = '(ninguno)' and WELLNAME IN (" & IDs & ")", BDVintage, adOpenStatic, adLockReadOnly, adCmdText
'  While Not RSVintage.EOF
'    If Not ExistePozo(RSVintage![WELLNAME]) Then
'      frmPozosFuturos.Barra.Panels("info") = "Generando pozo Well ID: " & RSVintage![WELLNAME]
'      RSLocal.Open "POZOS", BD, adOpenStatic, adLockOptimistic, adCmdTable
'      RSLocal.AddNew
'      RSLocal!WellID = RSVintage![WELLNAME]
'      RSLocal!Yacimiento = RSVintage!Area
'      RSLocal!XPOS94 = RSVintage!X
'      RSLocal!YPOS94 = RSVintage!Y
'      RSLocal!TD = RSVintage!TD
'      RSLocal!LANDOWNER = RSVintage!LANDOWNER
'      RSLocal!Ubicacion = "DRILLING INV."
'      RSLocal!OrdenUbicacion = ObtenerOrdenUbicacion("DRILLING INV.")
'      RSLocal!RIGORDER = 999
'      RSLocal!RIGORDERCHECKED = False
'      RSLocal.Update
'      'Grabo Historial
'
'      'vergpTEST
''      MsgBox "17"
'
'      BD.Execute "INSERT INTO HIST_POZOS (WELLID, Ubicacion, Equipo, MONOGRAFIA, ACUIFERO, Prognosis, Programa, Yacimiento, TIPOYACIMIENTO, Pozo, Prospect, FECHAPRIMERMONOGRAFIA, MONOGRAFIAS, FECHAULTIMAMONOGRAFIA, WELLINFORMED, INFORMEDBY, Definitiva, XPDC, YPDC, XPOS94, YPOS94, " & _
'             " FECHASOLICITUD, FECHAPRIORIDAD, DOCUMENTOAPREPARAR,FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO, FIRSTPROD, " & _
'             " TD, TOTDAYS, REMDAYS, Status, STARTDATE, ENDDATE, LANDOWNER, LANDOWNERPERMITDATE, " & _
'             " Consult, Type, FechaPedidoEtia, FECHAESPERADAETIA, " & _
'             " CONSULTANTRECOMENDATION, DMAPERMIT, Estado, " & _
'             " IDMANIFIESTO, FECHAMANIFIESTO, FECHAENTREGAEIAXCONSULTORAOXY, EIAPRESENTADO, FECHAENVIOACS, FECHAPRESENTACIONDMA, FECHAPRESENTACIONSMA, PAGOTASAADMINISTRATIVA, FECHAINFOCOMPLEMENTARIA, TECHNICALREPORT, TASAADMINISTRATIVA, TASACONTRALOR, Estudio, Adenda, FECHAINICIODIA, FECHAFINDIA,INFORMEDEAVANCEDEOBRA50PORCIENTO, INFORMEDEAVANCEDEOBRA100PORCIENTO, INFORMEEVALUACIONARQUEOLOGICO, " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA, TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA, TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA, TIEMPOENTREPRESENTACIONANTEDMAYVISITA, TIEMPOENTREVISITAYAPROBACIONFINALDEDMA, TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA, TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA, TIEMPOENTREENTREGAEIADMAYAPROBACION, TipoMovimiento, XID, XFECHA, XUSUARIO, RIGORDERCHECKED)" & _
'             " SELECT WELLID AS [Well ID], Ubicacion, Equipo, MONOGRAFIA AS [Monografía], ACUIFERO AS [Acuífero], Prognosis AS Prognosis, Programa AS Programa, Yacimiento, TIPOYACIMIENTO AS [Tipo Yacimiento], Pozo, Prospect, FECHAPRIMERMONOGRAFIA AS [Fecha Primera Monografia], MONOGRAFIAS AS [Monografías], FECHAULTIMAMONOGRAFIA AS [Fecha Ultima Monografía],  WELLINFORMED AS [Well Informed], INFORMEDBY AS [Informed By], Definitiva AS Definitiva, XPDC AS [X PDC], YPDC AS [Y PDC], XPOS94 AS [X POS94], YPOS94 AS [Y POS94], " & _
'             " FECHASOLICITUD AS [Fecha Solicitud], FECHAPRIORIDAD as [Fecha Prioridad], DOCUMENTOAPREPARAR AS [Documento a Preparar],FieldManifold, BatteryAssigned, FECHAENTREGADICTAMENTECNICO AS [Fecha Entrega Dictamen Tco], FIRSTPROD AS [First Prod], " & _
'             " TD, TOTDAYS AS [Tot Days], REMDAYS AS [Rem Days], Status, STARTDATE AS [Start Date], ENDDATE AS [End Date], LANDOWNER AS [Land Owner], LANDOWNERPERMITDATE AS [Land Owner Permit Date], " & _
'             " Consult, Type, FechaPedidoEtia AS [Fecha Pedido ETIA], FechaEsperadaEtia AS [Fecha Esperada ETIA], " & _
'             " CONSULTANTRECOMENDATION AS [Consultant Recomendation], DMAPERMIT AS [DMA Permit], Estado, " & _
'             " IDMANIFIESTO AS [ID Manifiesto], FECHAMANIFIESTO AS [Fecha Manifiesto], FECHAENTREGAEIAXCONSULTORAOXY AS [Fecha Entrega EIA Por Consultor a OXY], EIAPRESENTADO AS [EIA Presentado], FECHAENVIOACS AS [Fecha Envio ACS], FECHAPRESENTACIONDMA AS [Fecha Presentación DMA], FECHAPRESENTACIONSMA AS [Fecha Presentación SMA], PAGOTASAADMINISTRATIVA AS [Pago Tasa Administrativa], FECHAINFOCOMPLEMENTARIA AS [Fecha Info Complementaria], TECHNICALREPORT AS [Technical Report], TASAADMINISTRATIVA AS [Tasa Administrativa], TASACONTRALOR AS [Tasa Contralor], Estudio, Adenda, FECHAINICIODIA AS [Fecha Inicio DIA], FECHAFINDIA AS [Fecha Fin DIA], INFORMEDEAVANCEDEOBRA50PORCIENTO AS [Informe de Avance de Obra 50%], INFORMEDEAVANCEDEOBRA100PORCIENTO AS [Informe de Avance de Obra 100%], INFORMEEVALUACIONARQUEOLOGICO AS [Informe Evaluacion Arqueologico], " & _
'             " TIEMPOENTREPEDIDOETIAYRECEPCIONETIA AS [Tiempo Entre Pedido ETIA y Recepcion ETIA], TIEMPOENTREPRIMERMONOGRAFIAYRECEPCIONETIA AS [Tiempo Entre Primer Monografía y Recepción ETIA], TIEMPOENTRERECEPCIONETIAYPRESENTACIONANTEDMA AS [Tiempo Entre Recepcion ETIA y Presentación ante DMA], TIEMPOENTREPRESENTACIONANTEDMAYVISITA AS [Tiempo Entre Presentación Ante DMA y Visita], TIEMPOENTREVISITAYAPROBACIONFINALDEDMA AS [Tiempo Entre Visita y Aprobación Final De DMA], TIEMPOENTREPRESENTACIONDEPOZOYAPROBACIONFINALDMA AS [Tiempo Entre Presentación De Pozo y Aprobación Final DMA], TIEMPOENTREPRIMERMONOGRAFIAYPEDIDOEIA AS [Tiempo Entre Primer Monografia y Pedido EIA], TIEMPOENTREENTREGAEIADMAYAPROBACION AS [Tiempo Entre Entrega EIA a DMA y Aprobacion], 'A', IDPOZO, #" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "#,'Vintage Data', RIGORDERCHECKED FROM POZOS WHERE IDPOZO IN (" & RSLocal!IDPozo & ")"
'      RSLocal.Close
'    End If
'    RSVintage.MoveNext
'  Wend
'  RSVintage.Close
'  frmPozosFuturos.Barra.Panels("info") = "Generacion de pozos finalizada"
'
'ErrorHandler:
'  If Err.Number = 0 Then
'    BD.CommitTrans
'  Else
'    BD.RollbackTrans
'    ErrorHandler
'  End If
'End Sub


Private Function ExistePozo(WellID As String) As Boolean
  Dim rs As New Recordset
  On Error GoTo ErrorHandler
  rs.Open "SELECT WELLID FROM POZOS WHERE WELLID = '" & WellID & "'", BD, adOpenStatic, adLockReadOnly, adCmdText
  ExistePozo = (Not rs.EOF)
  rs.Close
  
ErrorHandler:
  If Err.Number <> 0 Then
    Err.Raise Err.Number
  End If
End Function

Public Sub DesconectarBDVintageData()
'Cierra la conexion con la base de datos del vintage data
  On Error GoTo ErrorHandler
  
  BDVintage.Close
  BDOXY.Close
  
ErrorHandler:
  ErrorHandler
End Sub

