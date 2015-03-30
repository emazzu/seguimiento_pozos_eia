Attribute VB_Name = "Enumeraciones"
Option Explicit

Public Enum enmTipoDeDato
  enmTipoDeDato_Entero = 3
  enmTipoDeDato_Decimal = 5
  enmTipoDeDato_Moneda = 6
  enmTipoDeDato_Texto = 202
  enmTipoDeDato_Memo = 203
  enmTipoDeDato_Fecha = 7
  enmTipoDeDato_Boolean = 11
  enmTipoDeDato_TipoYacimiento = 0
  enmTipoDeDato_Ubicacion = 1
End Enum

Public Enum enmColumnasLista 'Columnas de la lista del form de pozos
  colIDPozo
  colWellID
  colUbicacion
  colEquipo
  colYacimiento
  colTipoYacimiento
  colPozo
  colProspect
  colFechaPrimerMonografia
  colMonografias
  colFechaUltimaMonografia
  colWellInformed
  colInformedBy
  colDefinitiva
  colXPDC
  colYPDC
  colXPOS94
  colYPOS94
  colMonografia
  colAcuifero
  colPrognosis
  colPrograma
  colFechaSolicitud
  colFechaPrioridad
  colDocumentoAPreparar
  'vergpOK
  'colFechaEntregaPlanes8pts
  'colFechaEntregaPlanes13pts
  colFieldManifold
  colBatteryAssigned
  colFechaEntregaDictamenTecnico
  colFirstProd
  colTD
  colTotDays
  colRemDays
  colStatus
  colStartDate
  colEndDate
  colLandOwner
  colLandOwnerPermitDate
  colConsult
  colType
  colFechaPedidoETIA
  colFechaEsperadaETIA
  colImageInterpretationComments
  colConsultantRecomendation
  colSiteVisit
  colDMAFinalPermit
  colEstado
  colIDManifiesto
  colFechaManifiesto
  colFechaEntregaEIAxConsultoraAOXY
  colEIAPresentado
  colFechaEnvioACS
  colFechaPresentacionDMA
  colFechaPresentacionSMA
  colPagoTasaAdministrativa
  colFechaInfoComplementaria
  colTechnicalReport
  colTasaAdministrativa
  colTasaContralor
  colEstudio
  colAdenda
'  colAFE
'  colAFECerrado
  colFechainicioDIA
  colFechaFinDIA
  colInformeDeAvanceDeObra50PorCiento
  colInformeDeAvanceDeObra100PorCiento
  colInformeEvaluacionArqueologico
  colTiempoEntrePedidoETIAyRecepcionETIA
  colTiempoEntrePrimerMonografiaYRecepcionETIA
  colTiempoEntreRecepcionETIAYPresentacionAnteDMA
  colTIempoEntrePresentacionAnteDMAYVisita
  colTiempoEntreVisitaYAprobacionFinalDeDMA
  colTiempoEntrePresentacionDePozoYAprobacionFinalFMA
  colTiempoEntrePrimerMonografiaYPedidoEIA
  colTiempoEntreEntregaEIAaDMAyAprobacion
  colRigOrderChecked
End Enum

Public Enum enmEstados 'Estados posibles de los pozos
  enmEstadoTodos
  enmEstadoDrillingInv
  enmEstadoNuevosEstaqueos
  enmEstadoSuspendidos
  enmEstadoRigsScheduler
  enmEstadoPerforados
  enmEstadoNew
  enmEstadoGoneDone
  enmEstadoNuevos200
End Enum

'Columnas de las listas de comments del form de pozos
Public Enum enmColumnasComments
  colCommentsID
  colCommentsNumero
  colCommentsFecha
  colCommentsComment
  colCommentsAutor
  colCommentsNumeroACTA
End Enum
