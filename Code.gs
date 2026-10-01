/**
 * GENERADOR DE CARGUE NOVASOFT V2 
 * VERSIÓN DEFINITIVA: 118 Columnas + Cargo desde HR + codCargo
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Automatización Novasoft')
    .addItem('Generar Cargue por Fecha', 'abrirCalendario')
    .addToUi();
}

function abrirCalendario() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Calendario')
      .setWidth(320)
      .setHeight(230);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📅 Selector de Fecha Novasoft');
}

function generarCargueNovasoft(fechaSeleccionada) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zonaHoraria = ss.getSpreadsheetTimeZone(); 
  
  // 1. CONEXIONES
  const idHR = "1Rua64Q0MARrr_ChExM6U1KDigCG-wx8bx5SnwBMKMoQ";
  const sheetHR = SpreadsheetApp.openById(idHR).getSheetByName('HR COMPLETO');
  const idIngresos = "1a7iSi-8siRtDs-9qkpkk9rOUbf730KzbyzZV7UGzXyI";
  const sheetIngresos = SpreadsheetApp.openById(idIngresos).getSheetByName('ingresos 2026');
  const idSegSocial = "1tTxoUNKeQFon_FclKAgvYumjau-90FEG4QQcCS3lzcs";
  const appSegSocial = SpreadsheetApp.openById(idSegSocial);
  const sheetStaff = appSegSocial.getSheetByName('STAFF');
  const sheetAgentes = appSegSocial.getSheetByName('AGENTES');
  
  let sheetDatos = ss.getSheetByName('DATOS');
  if (!sheetDatos) sheetDatos = ss.insertSheet('DATOS');
  sheetDatos.clear();
  
  // ==========================================
  // 2. DICCIONARIOS DE REGLAS DE NEGOCIO
  // ==========================================
  
  const dicBancos = { "BANCO DE LA REPÚBLICA": "00", "BANCO DE BOGOTÁ": "01", "BANCO POPULAR": "02", "ITAÚ CORPBANCA COLOMBIA S.A.": "06", "BANCOLOMBIA S.A.": "07", "CITIBANK COLOMBIA": "09", "HSBC": "10", "GNB SUDAMERIS S.A.": "12", "BBVA COLOMBIA": "13", "BANCO DE CREDITO": "14", "SCOTIABANK COLPATRIA": "19", "BANCO DE OCCIDENTE": "23", "BCSC S.A.": "30", "BANCO CAJA SOCIAL - BCSC S.A.": "32", "INTERNACIONAL CIA FINANCIAMIEN": "33", "Banco Produban": "36", "BANCO AGRARIO DE COLOMBIA S.A.": "40", "BANCO DAVIVIENDA S.A.": "51", "BANCO AV VILLAS": "52", "BANCO W S.A.": "53", "BANCO COLMENA": "57", "BANCO PROCREDIT": "58", "BANCAMIA": "59", "BANCO PICHINCHA S.A.": "60", "BANCOOMEVA": "61", "CMR FALABELLA S.A.": "62", "BANCO FINANDINA S.A.": "63", "BANCO MULTIBANK S.A": "64", "BANCO SANTANDER COLOMBIA S.A.": "65", "BANCO COOPERATIVO COOPCENTRAL": "66", "BANCO COMPARTIR S.A": "67", "BANCO SERFINANZA S.A": "69", "LULO BANK S.A.": "70", "NEQUI": "74", "DAVIPLATA": "75", "FINANCIERA JURISCOOP": "76", "COOP FINANCIERA DE ANTIOQUIA": "77", "COOTRAFA COOP FINANCIERA": "78", "CONFIAR COOPERATIVA FINANCIERA": "79", "COLTEFINANCIERA S.A": "80", "EC PACIFIC": "81", "Banco Correval": "90", "Credicorp Capital": "A1", "RAPPIPAY": "15", "Ualá": "16", "NU BANK": "82", "NU": "82", "Dale": "17" };
  const dicSucursales = { "Accounting and Payroll Admin": "35", "AIRBNB Admin": "35", "AIRBNB Agent Backoffice": "35", "AIRBNB Agent Voice": "35", "AIRBNB Agente Email": "35", "BOOKING Admin": "35", "BOOKING Agente Backoffice": "35", "BOOKING Agente Email": "35", "BOOKING Agente Voice": "35", "Campus Operations Admin": "35", "Campus Recruiting Operations Admin": "35", "Corporate Recruiting Admin": "35", "Customer Experience Admin": "35", "EX Admin": "35", "Facilities Admin": "35", "Global Operations Admin": "35", "GNC Admin": "35", "GNC Agente Backoffice": "35", "GNC Agente Email": "35", "GNC Agente Voice": "35", "Human Resources Admin": "35", "IT Operations Admin": "35", "LEVIS Admin": "35", "LEVIS Agente Backoffice": "35", "LEVIS Agente Email": "35", "LEVIS Agente Voice": "35", "LOREAL Admin": "35", "LOREAL Agente Backoffice": "35", "LOREAL Agente Email": "35", "LOREAL Agente Voice": "35", "Marketing Admin": "35", "Operations Leadership Admin": "35", "OPTAVIA Admin": "35", "OPTAVIA Agente Backoffice": "35", "OPTAVIA Agente Email": "35", "OPTAVIA Agente Voice": "35", "SEPHORA Admin": "20", "SEPHORA Agente Backoffice": "20", "SEPHORA Agente Email": "20", "SEPHORA Agente Voice": "20", "SPARK Agente Backoffice": "35", "SPARK Agente Chat": "35", "SPARK Chat Admin": "35", "SPARK Agente Email": "35", "SPARK Agente Voice": "35", "SPARK Email Admin": "35", "SPARK Voice Admin": "35", "TCP Admin": "35", "TCP Agente Backoffice": "35", "TCP Agente Email": "35", "TCP Agente Voice": "35", "Training Admin": "35", "VROOM Admin": "35", "VROOM Agente Backoffice": "35", "VROOM Agente Email": "35", "VROOM Agente Voice": "35", "WEBER Admin": "35", "WEBER Agente Backoffice": "35", "WEBER Agente Email": "35", "WEBER Agente Voice": "35", "MEJURI Admin": "35", "MEJURI Agente Backoffice": "35", "RED ROBIN Admin": "35", "RED ROBIN Agente Backoffice": "35", "Marigold Admin": "35", "Marigold Agente Voice": "35", "Noom Admin": "35", "Noom Agente Voice": "35", "CarGurus Admin": "35", "CarGurus Agente Voice": "35", "NYT Admin": "35", "NYT Agente Voice": "35", "Newell Admin": "35", "Newell Agente Voice": "35", "StubHub Admin": "35", "StubHub Agente Voice": "35", "HCA Servicing Admin": "35", "HCA Servicing Agente Voice": "35", "Workforce": "35", "EnFin Agente Voice": "35", "EnFin Admin": "35", "Campus Global": "35", "Pandora Admin": "35", "Pandora Agente Voice": "35", "PetSmart Admin": "35", "PetSmart Agente Voice": "35", "NCAA Admin": "35", "NCAA Agente Voice": "35", "MLB Admin": "35", "MLB Agente Voice": "35", "Aritzia Agente Voice": "35", "Aritzia": "35" };
  const dicGastos = { "Accounting and Payroll Admin": "52", "AIRBNB Admin": "51", "AIRBNB Agent Backoffice": "61", "AIRBNB Agent Voice": "61", "AIRBNB Agente Email": "61", "BOOKING Admin": "51", "BOOKING Agente Backoffice": "61", "BOOKING Agente Email": "61", "BOOKING Agente Voice": "61", "Campus Operations Admin": "52", "Campus Recruiting Operations Admin": "52", "Corporate Recruiting Admin": "52", "Customer Experience Admin": "51", "EX Admin": "52", "Facilities Admin": "52", "Global Operations Admin": "52", "GNC Admin": "51", "GNC Agente Backoffice": "61", "GNC Agente Email": "61", "GNC Agente Voice": "61", "Human Resources Admin": "52", "IT Operations Admin": "52", "LEVIS Admin": "51", "LEVIS Agente Backoffice": "61", "LEVIS Agente Email": "61", "LEVIS Agente Voice": "61", "LOREAL Admin": "51", "LOREAL Agente Backoffice": "61", "LOREAL Agente Email": "61", "LOREAL Agente Voice": "61", "Marketing Admin": "52", "Operations Leadership Admin": "52", "OPTAVIA Admin": "51", "OPTAVIA Agente Backoffice": "61", "OPTAVIA Agente Email": "61", "OPTAVIA Agente Voice": "61", "SEPHORA Admin": "51", "SEPHORA Agente Backoffice": "61", "SEPHORA Agente Email": "61", "SEPHORA Agente Voice": "61", "SPARK Agente Backoffice": "61", "SPARK Agente Chat": "61", "SPARK Chat Admin": "51", "SPARK Agente Email": "61", "SPARK Agente Voice": "61", "SPARK Email Admin": "51", "SPARK Voice Admin": "51", "TCP Admin": "51", "TCP Agente Backoffice": "61", "TCP Agente Email": "61", "TCP Agente Voice": "61", "Training Admin": "52", "VROOM Admin": "51", "VROOM Agente Backoffice": "61", "VROOM Agente Email": "61", "VROOM Agente Voice": "61", "WEBER Admin": "51", "WEBER Agente Backoffice": "61", "WEBER Agente Email": "61", "WEBER Agente Voice": "61", "MEJURI Admin": "51", "MEJURI Agente Backoffice": "61", "RED ROBIN Admin": "51", "RED ROBIN Agente Backoffice": "61", "Marigold Admin": "51", "Marigold Agente Voice": "61", "Noom Admin": "51", "Noom Agente Voice": "61", "CarGurus Admin": "51", "CarGurus Agente Voice": "61", "NYT Admin": "51", "NYT Agente Voice": "61", "Newell Admin": "51", "Newell Agente Voice": "61", "StubHub Admin": "51", "StubHub Agente Voice": "61", "HCA Servicing Admin": "51", "HCA Servicing Agente Voice": "61", "Workforce": "51", "EnFin Agente Voice": "61", "EnFin Admin": "51", "Campus Global": "51", "Pandora Admin": "51", "Pandora Agente Voice": "61", "PetSmart Admin": "51", "PetSmart Agente Voice": "61", "NCAA Admin": "51", "NCAA Agente Voice": "61", "MLB Admin": "51", "MLB Agente Voice": "61", "Aritzia Agente Voice": "61", "Aritzia": "61" };
  const dicFondosPension = { "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES": "108", "CAJA DE AUXILIOS Y DE PRESTACIONES DE ACDAC": "105", "COLFONDOS": "104", "FONDO DE PREVISIÓN SOCIAL DEL CONGRESO": "106", "NO APLICA (APRENDICES)": "0", "NO APLICA PENSION": "199", "SKANDIA": "102", "PENSIONES DE ANTIOQUIA": "107", "PORVENIR": "101", "PROTECCIÓN": "100", "OLD MUTUAL FONDO DE PENSIONES OBLIGATORIAS": "102" };
  const dicFondosSalud = { "ALIANSALUD EPS": "200", "ASMET SALUD EPS": "224", "CAJA DE PREVISIÓN SOCIAL DE CASANARE - CAPRESOCA E.P.S.": "231", "CAJA DE PREVISIÓN SOCIAL DE COMUNICACIONES - CAPRECOM": "230", "PROTEGER EPS S.A.S": "225", "CAPITAL SALUD": "218", "COMFACUNDI": "220", "COMFAMILIAR HUILA": "227", "COMFENALCO VALLE EPS": "206", "COMPENSAR ENTIDAD PROMOTORA DE SALUD": "204", "COMFAORIENTE EPS-S": "234", "COOSALUD EPS": "219", "E.P.S SANITAS": "203", "EMPRESA COOPERATIVA SOLIDARIA DE SALUD “ECOOPSOS”": "223", "EMPRESAS PÚBLICAS DE MEDELLÍN DEPARTAMENTO MÉDICO": "216", "EMSSANAR": "221", "EPS SURA": "205", "FAMISANAR": "209", "FONDO DE PASIVO SOCIAL DE FERROCARRILES": "217", "FONDO DE SOLIDARIDAD Y GARANTÍA FOSYGA": "215", "Fundación Salud Mía EPS": "229", "MUTUAL SER E.S.S.": "222", "NO APLICA": "299", "NUEVA EPS": "214", "SALUD TOTAL S.A.": "201", "SALUDVIDA S.A EPS": "213", "SAVIA SALUD EPS": "228", "SERVICIO OCCIDENTAL DE SALUD S.A. S.O.S.": "210", "SALUD BOLÍVAR EPS S.A.S": "233", "EPS FAMILIAR DE COLOMBIA S.A.S.": "236" };
  const dicFondosCesantias = { "CESANTIAS COLFONDOS": "503", "CESANTIAS PORVENIR": "502", "CESANTIAS PROTECCIÓN": "501", "CESANTIAS SKANDIA": "504", "FONDO NACIONAL DEL AHORRO - FNA": "500", "NO APLICA": "599" };
  
  const dicClasificador1 = { "Auxiliar de Recursos Humanos": "1002", "Especialista de Recursos Humanos": "1002", "Generalista de Recursos Humanos": "1002", "Gerente de Recursos Humanos": "1002", "HR Business Partner": "1002", "Supervisor": "1011", "Analista de Recursos Humanos PP": "1002", "Employee Services Business Partner": "1002", "Director de Reclutamiento": "1002", "Lider de Reclutamiento": "1002", "Lider de Sourcing": "1002", "Reclutador": "1002", "Sourcer": "1002", "Analista de Reclutamiento": "1002", "Reclutador Front Desk": "1002", "Analista de Facilities": "1002", "Analista de SST": "1002", "Aprendiz Sena": "1004", "Aprendiz Universitario": "1004", "Asistente": "1002", "Coodinador de Facilities": "1002", "Coordinador de Seguridad y Salud en el Trabajo": "1002", "Director de Campus": "1002", "Generalista de Seguridad y Salud en el Trabajo": "1002", "Gerente de Capacitacion": "1002", "Gerente de Facilities": "1002", "Gerente Seguridad y Salud en el Trabajo": "1002", "Lider de Concierge": "1002", "VP Campus Operations LATAM / Country Manager": "1002", "Analista de Negocios": "1013", "Gerente Jr Employee Experience": "1002", "Oficial de Cumplimiento y Seguridad": "1002", "Asistente de Facilities": "1002", "Analista Senior Employee Experience": "1002", "Brand & Marketing Coordinator": "1002", "Copywriter": "1001", "Gerente de Comunicaciones": "1002", "Director de Comunicaciones": "1002", "Content Writter Jr": "1002", "Generalista IT": "1002", "Generalista Senior IT": "1002", "Gerente Jr IT": "1002", "Analista de Reporteria": "1014", "Concierge": "1015", "Analista Financiero": "1013", "Gerente de Analitica de Negocios Senior": "1013", "Ingeniero de Datos": "1013", "Mentor de Desempeño": "1013", "Analista de Tiempo Real": "1014", "Gerente de Reporteria": "1009", "Gerente de Workforce": "1014", "Planificador de Recursos": "1014", "Especialista de Horarios": "1014", "Supervisor de Workforce": "1014", "Gerente de Workforce Senior": "1014", "Analista de Datos Senior": "1014", "Supervisor de Reporteria": "1014", "Agente Telefonico": "1011", "Agente Telefonico II": "1001", "CSR2": "1008", "CSR2 PP": "1008", "Interinato CSR2": "1008", "Supervisor PP": "1001", "Analista de Calidad": "1006", "Analista de Negocios Senior": "1013", "Capacitador": "1003", "Gerente de Capacitacion y Calidad": "1003", "Gerente de Operaciones": "1007", "Gerente de Operaciones Senior": "1007", "Interinato Analista de Calidad": "1006", "Interinato Capacitador": "1003", "Interinato Supervisor": "1011", "Senior Supervisor": "1011", "SME": "1008", "Supervisor de Calidad": "1006", "Supervisor de Entrenamiento": "1003", "Supervisor de Capacitación y Calidad PP": "1006", "Director de Operaciones": "1012", "Director de Operaciones Senior": "1012", "Director Senior de Reporteria": "1009", "Director de Workforce": "1012", "Director de Workforce Senior": "1012", "Analista de Contabilidad": "1002", "Analista de Finanzas": "1002", "Analista de Nomina": "1002", "Auxiliar de Nomina": "1002", "Controller": "1002", "Coordinador Contabilidad": "1002", "Coordinador de Nomina": "1002", "Gerente de Nomina": "1002", "Interinato Coordinador de Nomina": "1002", "Asistente de Nomina": "1002", "Reclutador Corporativo": "1002", "Business Development representative": "1002", "Analista de Datos": "1014", "Auxiliar": "1002" };
  
  const dicCodCargo = {
    "Agente Telefonico": "001100",
    "Analista de Calidad": "001020",
    "Analista de Contabilidad": "010011",
    "Analista de desempeño": "001021",
    "Analista de Negocios": "001023",
    "Analista de Nomina": "010021",
    "Analista de Recursos Humanos": "030012",
    "Analista de Reporteria": "080010",
    "Analista de Tiempo Real": "001022",
    "Aprendiz etapa lectiva": "200001",
    "Aprendiz etapa productiva": "200002",
    "Auxiliar de Facilities": "100004",
    "Auxiliar de IT": "060015",
    "Auxiliar de Recursos Humanos": "030015",
    "Auxiliar de Nomina": "010022",
    "Capacitador": "050011",
    "Controller": "010001",
    "Coodinador de Recursos Humanos": "030010",
    "Coordinador Contabilidad": "010010",
    "Coordinador de Facilities": "100003",
    "Coordinador de Nómina": "010020",
    "Coordinador de Reclutamiento": "020010",
    "Coordinador de Seguridad y Salud en el Trabajo": "100010",
    "Coordinador Entrenamiento": "050010",
    "Director de Operaciones": "090001",
    "Director de Reclutamiento": "020002",
    "Especialista en Comunicaciones": "110005",
    "Front Desk Recuiter": "020012",
    "Generalista de Recursos Humanos": "030011",
    "Generalista IT": "060011",
    "Gerente de Capacitación": "050001",
    "Gerente de Capacitacion y Calidad": "001002",
    "Gerente de HSE": "014001",
    "Gerente de Instalaciones": "100002",
    "Gerente de Marketing": "110001",
    "Gerente de Nómina": "010005",
    "Gerente de Operaciones": "001005",
    "Gerente de Operaciones Senior": "001001",
    "Gerente de Reclutamiento": "020001",
    "Gerente de Recursos Humanos": "030001",
    "Gerente de Reporteria": "080002",
    "Aprendiz Sena": "200002",
    "Gerente Employee Experience": "070001",
    "Gerente Facilities": "100001",
    "Gerente Ops Support": "080001",
    "Gerente País": "040001",
    "Lider de Reclutamiento": "020003",
    "New Hire Champion": "001011",
    "NO APLICA": "000000",
    "Reclutador": "020011",
    "Reclutador Corporativo": "020015",
    "Source lead": "020004",
    "Sourcer": "020013",
    "Sourcer - Recruiter": "020014",
    "Supervisor": "001010",
    "Supervisor de Calidad": "001019",
    "Supervisor de Entrenamiento": "050012",
    "Supervisor de IT": "060010",
    "Supervisor Senior": "001007",
    "Supervisor Senior IT": "060005",
    "Director Campus": "040002",
    "Director Senior de Reporteria": "080000",
    "Generalista SST": "014002",
    "Supervisor de Soporte": "010006",
    "Especialista de Recursos Humanos": "030013",
    "Aprendiz Universitario": "200002",
    "Asistente": "110025",
    "Analista de Finanzas": "010011",
    "Analista de SST": "030020",
    "Analista de Negocios Senior": "001024",
    "Auxiliar de SST": "014002",
    "Brand & Marketing Coordinator": "110007",
    "Senior Manager, Demand Generation": "030026",
    "Gerente de Facilities": "100001",
    "Copywriter": "110008",
    "Director de Reporteria": "080003",
    "Forecasting Analyst": "110010",
    "Gerente de Reporteria Senior": "110011",
    "Concierge": "030014",
    "Scheduling Specialist": "110013",
    "Resource Planner": "110015",
    "Senior Manager Business Analytics": "110016",
    "Gerente de Workforce": "110031",
    "Analista de Facilities": "110020",
    "Gerente Senior de Prevencion de Fraude": "110019",
    "HR Business Partner": "110026",
    "Analista de Datos": "110028",
    "Director de Comunicaciones": "110029",
    "Especialista de Horarios": "110030",
    "Director de Workforce": "110017",
    "Ingeniero de Datos": "110044",
    "Planificador de Recursos": "110035",
    "Gerente de Comunicaciones": "110043",
    "Supervisor de juegos y espec": "110049",
    "Supervisor de Reporteria": "110047"
  };

  const dicClasificador2 = { "Accounting and Payroll Admin": "0", "AIRBNB Admin": "0", "AIRBNB Agent Backoffice": "6", "AIRBNB Agent Voice": "2", "AIRBNB Agente Email": "3", "BOOKING Admin": "0", "BOOKING Agente Backoffice": "6", "BOOKING Agente Email": "3", "BOOKING Agente Voice": "2", "Campus Operations Admin": "0", "Campus Recruiting Operations Admin": "0", "Corporate Recruiting Admin": "0", "Customer Experience Admin": "0", "EX Admin": "0", "Facilities Admin": "0", "Global Operations Admin": "0", "GNC Admin": "0", "GNC Agente Backoffice": "6", "GNC Agente Email": "3", "GNC Agente Voice": "2", "Human Resources Admin": "0", "IT Operations Admin": "0", "LEVIS Admin": "0", "LEVIS Agente Backoffice": "6", "LEVIS Agente Email": "3", "LEVIS Agente Voice": "2", "LOREAL Admin": "0", "LOREAL Agente Backoffice": "6", "LOREAL Agente Email": "3", "LOREAL Agente Voice": "2", "Marketing Admin": "0", "Operations Leadership Admin": "0", "OPTAVIA Admin": "0", "OPTAVIA Agente Backoffice": "6", "OPTAVIA Agente Email": "3", "OPTAVIA Agente Voice": "2", "SEPHORA Admin": "0", "SEPHORA Agente Backoffice": "6", "SEPHORA Agente Email": "3", "SEPHORA Agente Voice": "2", "SPARK Agente Backoffice": "6", "SPARK Agente Chat": "1", "SPARK Chat Admin": "1", "SPARK Agente Email": "3", "SPARK Agente Voice": "2", "SPARK Email Admin": "0", "SPARK Voice Admin": "0", "TCP Admin": "0", "TCP Agente Backoffice": "6", "TCP Agente Email": "3", "TCP Agente Voice": "2", "Training Admin": "0", "VROOM Admin": "0", "VROOM Agente Backoffice": "6", "VROOM Agente Email": "3", "VROOM Agente Voice": "2", "WEBER Admin": "0", "WEBER Agente Backoffice": "6", "WEBER Agente Email": "3", "WEBER Agente Voice": "2", "MEJURI Admin": "0", "MEJURI Agente Backoffice": "6", "RED ROBIN Admin": "0", "RED ROBIN Agente Backoffice": "6", "Marigold Admin": "0", "Marigold Agente Voice": "6", "Noom Admin": "0", "Noom Agente Voice": "6", "CarGurus Admin": "0", "CarGurus Agente Voice": "6", "NYT Admin": "0", "NYT Agente Voice": "6", "Newell Admin": "0", "Newell Agente Voice": "0", "StubHub Admin": "0", "StubHub Agente Voice": "0", "HCA Servicing Admin": "0", "HCA Servicing Agente Voice": "0", "Workforce": "0", "EnFin Agente Voice": "0", "EnFin Admin": "0", "Campus Global": "0", "Pandora Admin": "0", "Pandora Agente Voice": "0", "PetSmart Admin": "0", "PetSmart Agente Voice": "0", "NCAA Admin": "0", "NCAA Agente Voice": "0", "MLB Admin": "0", "MLB Agente Voice": "0", "Aritzia Agente Voice": "2", "Aritzia": "2" };
  const dicClasificador7 = { "Accounting and Payroll Admin": "8", "AIRBNB Admin": "520", "AIRBNB Agent Backoffice": "519", "AIRBNB Agent Voice": "519", "AIRBNB Agente Email": "519", "BOOKING Admin": "520", "BOOKING Agente Backoffice": "519", "BOOKING Agente Email": "519", "BOOKING Agente Voice": "519", "Campus Operations Admin": "31", "Campus Recruiting Operations Admin": "27", "Corporate Recruiting Admin": "28", "Customer Experience Admin": "31", "EX Admin": "31", "Facilities Admin": "31", "Global Operations Admin": "107", "GNC Admin": "520", "GNC Agente Backoffice": "519", "GNC Agente Email": "519", "GNC Agente Voice": "519", "Human Resources Admin": "14", "IT Operations Admin": "45", "LEVIS Admin": "520", "LEVIS Agente Backoffice": "519", "LEVIS Agente Email": "519", "LEVIS Agente Voice": "519", "LOREAL Admin": "520", "LOREAL Agente Backoffice": "519", "LOREAL Agente Email": "519", "LOREAL Agente Voice": "519", "Marketing Admin": "38", "Operations Leadership Admin": "67", "OPTAVIA Admin": "520", "OPTAVIA Agente Backoffice": "519", "OPTAVIA Agente Email": "519", "OPTAVIA Agente Voice": "519", "SEPHORA Admin": "520", "SEPHORA Agente Backoffice": "519", "SEPHORA Agente Email": "519", "SEPHORA Agente Voice": "519", "SPARK Agente Backoffice": "519", "SPARK Agente Chat": "519", "SPARK Chat Admin": "520", "SPARK Agente Email": "519", "SPARK Agente Voice": "519", "SPARK Email Admin": "520", "SPARK Voice Admin": "520", "TCP Admin": "520", "TCP Agente Backoffice": "519", "TCP Agente Email": "519", "TCP Agente Voice": "519", "Training Admin": "19", "VROOM Admin": "520", "VROOM Agente Backoffice": "519", "VROOM Agente Email": "519", "VROOM Agente Voice": "519", "WEBER Admin": "520", "WEBER Agente Backoffice": "519", "WEBER Agente Email": "519", "WEBER Agente Voice": "519", "MEJURI Admin": "520", "MEJURI Agente Backoffice": "519", "RED ROBIN Admin": "520", "RED ROBIN Agente Backoffice": "519", "Marigold Admin": "520", "Marigold Agente Voice": "519", "Noom Admin": "520", "Noom Agente Voice": "519", "CarGurus Admin": "520", "CarGurus Agente Voice": "519", "NYT Admin": "520", "NYT Agente Voice": "519", "Newell Admin": "520", "Newell Agente Voice": "519", "StubHub Admin": "520", "StubHub Agente Voice": "519", "HCA Servicing Admin": "520", "HCA Servicing Agente Voice": "519", "Workforce": "518", "EnFin Agente Voice": "519", "EnFin Admin": "520", "Campus Global": "10", "Pandora Admin": "520", "Pandora Agente Voice": "519", "PetSmart Admin": "520", "PetSmart Agente Voice": "519", "NCAA Admin": "520", "NCAA Agente Voice": "519", "MLB Admin": "520", "MLB Agente Voice": "519", "Aritzia Agente Voice": "519", "Aritzia": "519" };
  const dicCentroTrabajo = { "Accounting and Payroll Admin": "1", "AIRBNB Admin": "1", "AIRBNB Agent Backoffice": "1", "AIRBNB Agent Voice": "1", "Aritzia Agente Voice": "1", "Aritzia": "1" };
  const dicCodigoArea = { "Aritzia Agente Voice": "0", "Aritzia": "0" };

  const dicFondoCCF = { "BOGOTA": "CCF21", "BUCARAMANGA": "CCF39", "CUNDINAMARCA": "CCF21", "BARRANQUILLA": "CCF07", "CALI": "CCF57", "MEDELLIN": "CCF03", "TUNJA": "CCF10" };
  const dicPaises = { "COLOMBIA": "169", "VENEZUELA": "850", "ECUADOR": "218", "ESTADOS UNIDOS": "249" };

  function extraerPais(textoLugar) {
    if (!textoLugar) return "169"; 
    let textoUP = String(textoLugar).toUpperCase();
    for (let pais in dicPaises) {
      if (textoUP.includes(pais)) return dicPaises[pais];
    }
    return "169";
  }

  // 3. LEER DATA DE HOJAS
  const dataHR = sheetHR.getDataRange().getValues();
  const dataIngresosDisplay = sheetIngresos.getDataRange().getDisplayValues();
  const dataIngresosValores = sheetIngresos.getDataRange().getValues(); 
  const dataStaff = sheetStaff.getDataRange().getValues();
  const dataAgentes = sheetAgentes.getDataRange().getValues();
  
  // 4. MAPEO DINÁMICO DE COLUMNAS PARA INGRESOS
  let idxIng = { fecha: 0, cde: 1, nombre: 2, cc: 4, banco: 21, contrato: 14, salario: 16, elec: 23, cuenta: 20, rh: 7, tipoDoc: 9 };
  for(let r = 0; r < Math.min(dataIngresosValores.length, 10); r++) {
    for(let c = 0; c < dataIngresosValores[r].length; c++) {
      let head = String(dataIngresosValores[r][c]).toUpperCase().trim();
      if(head==="FECHA" || head.includes("START DATE")) idxIng.fecha = c;
      if(head==="CDE" || head==="EID" || head.includes("NATIONAL ID")) idxIng.cde = c;
      if(head==="DEF CENTRO COSTO" || head.includes("CLASIFICADOR 1")) idxIng.cc = c;
      if(head.includes("ENTIDAD BANCARIA")) idxIng.banco = c;
      if(head==="SALARIO") idxIng.salario = c;
      if(head.includes("TIPO CONTRATO")) idxIng.contrato = c;
      if(head.includes("NO. CUENTA") || head === "CUENTA") idxIng.cuenta = c;
      if(head.includes("CIUDAD EXPEDICCION")) idxIng.ciudadExp = c;
      if(head.includes("CIUDAD NACIMIENTO")) idxIng.ciudadNac = c;
      if(head.includes("FECHA FIN") || head.includes("FINALIZACION")) idxIng.fechaFin = c;
      if(head.includes("TYPE CLASIF 3") || head === "CLASIF 3") idxIng.clasif3 = c;
      if(head.includes("B. PRODUCTIVIDAD")) idxIng.prod = c;
      if(head.includes("B. ATTENDANCE")) idxIng.att = c;
      if(head.includes("SITE") && !head.includes("ID")) idxIng.site = c;
      if(head==="RH" || head.includes("GRUPO SANG")) idxIng.rh = c;
      if(head.includes("PAGO ELECTRONICO") || head.includes("ID PAGO")) idxIng.elec = c;
      if(head.includes("T.DOCUMENTO") || head.includes("TIPO DOC")) idxIng.tipoDoc = c;
    }
    if (dataIngresosValores[r][idxIng.cde] !== undefined && String(dataIngresosValores[r][idxIng.cde]).trim() !== "") { 
      idxIng.headerRow = r; break; 
    }
  }
  
  // MAPEO DE HR
  // *AQUÍ SE EXTRAE EL CARGO DESDE RRHH EN LUGAR DE INGRESOS*
  let idxHR = { eid: 0, cedula: 23, ciudad: 45, legal: 49, horas: 41, cargo: 9 };
  for(let c=0; c < dataHR[0].length; c++) {
    let head = String(dataHR[0][c]).toUpperCase().trim();
    if(head==="EID" || head==="CODIGO" || head==="ID") idxHR.eid = c; 
    if(head.includes("NATIONAL ID") || head.includes("CEDULA")) idxHR.cedula = c;
    if(head.includes("HIRING CITY") || head.includes("CIUD UBIC")) idxHR.ciudad = c;
    if(head.includes("LEGAL ENTITY") || head==="EMPRESA") idxHR.legal = c;
    if(head.includes("HOURS") || head==="HORAS") idxHR.horas = c;
    if(head==="TITLE" || head==="CARGO" || head==="PUESTO") idxHR.cargo = c;
  }

  // 5. CRUCE DE CÉDULAS
  let eidToCedula = {};
  for (let i = 1; i < dataHR.length; i++) {
    let e = String(dataHR[i][idxHR.eid]).trim();  
    let c = String(dataHR[i][idxHR.cedula]).trim(); 
    if (e && c) eidToCedula[e] = c;
  }

  let ingresosPorCedula = {};
  let fechaRealPorCedula = {}; 
  
  let inicioIngresos = idxIng.headerRow !== undefined ? idxIng.headerRow + 1 : 1;
  for (let i = inicioIngresos; i < dataIngresosDisplay.length; i++) {
    let idStr = String(dataIngresosDisplay[i][idxIng.cde] || "").trim(); 
    if (idStr === "") continue;
    let cedulaDefinitiva = eidToCedula[idStr] || idStr; 
    ingresosPorCedula[cedulaDefinitiva] = dataIngresosDisplay[i];
    fechaRealPorCedula[cedulaDefinitiva] = dataIngresosValores[i][idxIng.fecha]; 
  }
  
  let fondosPorCedula = {};
  for (let i = 1; i < dataStaff.length; i++) {
    let cedula = String(dataStaff[i][4]).trim();
    fondosPorCedula[cedula] = { cesantias: dataStaff[i][6], eps: dataStaff[i][7], pension: dataStaff[i][8] };
  }
  for (let i = 1; i < dataAgentes.length; i++) {
    let cedula = String(dataAgentes[i][4]).trim();
    fondosPorCedula[cedula] = { cesantias: dataAgentes[i][6], eps: dataAgentes[i][7], pension: dataAgentes[i][8] };
  }
  
  // 6. ENCABEZADOS (118)
  const encabezados = ['codEmp', 'nroIdentificacion', 'tipoIdentificacion', 'primerApellido', 'primerNombre', 'fechaNacimiento', 'sexoEmp', 'claseLibreta', 'estadocivil', 'nacionalidad', 'direccionRes', 'email', 'celular', 'fechaIngreso', 'tipoPago', 'banco', 'cuentaBanco', 'tipoLIquida', 'regimenSal', 'claseSalario', 'compania', 'Sucursal', 'centroCosto', 'clasificador1', 'clasificador2', 'clasificador3', 'clasificador4', 'clasificador5', 'clasificador6', 'clasificador7', 'codCargo', 'tipoContrato', 'fondoRiesgos', 'porRiesgos', 'fondoPension', 'fondoSalud', 'fondoCCF', 'fondoCesant', 'metReten', 'salario', 'diasVacacionAno', 'sabadoHabil', 'promedioSalud', 'cuentaGasto', 'variable', 'porDias', 'modoLiquidacion', 'paisExpIdentif', 'ciudadExpIdent', 'paisNacimiento', 'ciudadNacimiento', 'paisRes', 'ciudadRes', 'paisLab', 'ciuLabor', 'paisContrato', 'ciuContrato', 'pensionado', 'pensxEmpresa', 'tipCotizante', 'subtipoCotizante', 'extranjero', 'resideExtranjero', 'ley1393', 'contratista', 'ley1450', 'horasMes', 'nroContrato', 'sucursalSS', 'cuentaGastoNiif', 'distribucionCCosto', 'distriCCostoNiif', 'deducibleSalud', 'deducibleVivienda', 'dependientes', 'nroPersoCargo', 'valorHora', 'codPagoElectronico', 'PagarDia31Vac', 'Localidad', 'CodBarrio', 'NroPasaporte', 'PaisEmisorPasaporte', 'codigoAlterno', 'segundoApellido', 'segundoNombre', 'numLibreta', 'distritoLibreta', 'grupoSang', 'factorRh', 'telefonoRes', 'barrio', 'porcReten', 'convenio', 'fechaFinContrato', 'codCentDeTrabajo', 'codigoArea', 'sueldoAntesFlex', 'porcenFlex', 'estatura', 'peso', 'emailAlterno', 'sucursalConvenio', 'ccostoConvenio', 'clasificador1Conv', 'clasificador2Conv', 'clasificador3Conv', 'clasificador4Conv', 'clasificador5Conv', 'clasificador6Conv', 'clasificador7Conv', 'clasificador8Conv', 'lider', 'FechaExpdocId', 'FechaExpPasaporte', 'ContratacionExterna', 'TipoTrabajo', 'IndTransicion'];
  
  let resultadoFinal = [encabezados];
  const indicesCeros = [41, 42, 44, 45, 57, 58, 60, 61, 62, 64, 65, 68, 71, 72, 73, 74, 75, 78, 79, 80, 81, 82, 87, 92, 93, 97, 98, 99, 100, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 117];
  
  // 7. PROCESAMOS DATA
  for (let i = 1; i < dataHR.length; i++) {
    let rowHR = dataHR[i];
    let cedula = String(rowHR[idxHR.cedula]).trim();
    if (!cedula || cedula === "") continue; 
    
    let ingresos = ingresosPorCedula[cedula];
    if (!ingresos) continue; 
    
    let fechaCelda = fechaRealPorCedula[cedula];
    let matchFecha = false;
    
    if (fechaCelda) {
      if (fechaCelda instanceof Date) {
        let f1 = Utilities.formatDate(fechaCelda, "GMT", "yyyy-MM-dd");
        let f2 = Utilities.formatDate(fechaCelda, zonaHoraria, "yyyy-MM-dd");
        if (f1 === fechaSeleccionada || f2 === fechaSeleccionada) matchFecha = true;
      } else {
        let txt = String(fechaCelda).trim();
        if (txt.includes(fechaSeleccionada)) matchFecha = true;
      }
    }
    
    if (!matchFecha) continue; 
    
    let fondos = fondosPorCedula[cedula] || { eps: '', pension: '', cesantias: '' };
    
    let centroCosto = idxIng.cc !== undefined ? String(ingresos[idxIng.cc] || "").trim() : ""; 
    let tipoContrato = idxIng.contrato !== undefined ? String(ingresos[idxIng.contrato]).trim() : ""; 
    // AQUÍ ESTÁ EL CAMBIO MAESTRO: Tomar el cargo desde HR
    let cargo = rowHR[idxHR.cargo] ? String(rowHR[idxHR.cargo]).trim() : ""; 
    let bancoTexto = idxIng.banco !== undefined ? String(ingresos[idxIng.banco]).trim() : "";
    let salario = idxIng.salario !== undefined ? Number(ingresos[idxIng.salario]) : 0; 
    let idPagoElectronico = idxIng.elec !== undefined ? String(ingresos[idxIng.elec]).trim() : ""; 
    
    let horasMesCalculado = (tipoContrato === "10" || tipoContrato === "11" || cargo.toUpperCase() !== "AGENTE TELEFONICO") ? 210 : (rowHR[idxHR.horas] || 210);
    let valorHoraCalc = horasMesCalculado > 0 ? Math.round(salario / horasMesCalculado) : 0;
    
    let filaSalida = new Array(118).fill(""); 
    indicesCeros.forEach(idx => filaSalida[idx] = "0");
    
    filaSalida[0] = rowHR[0]; 
    filaSalida[1] = cedula;   
    
    let tipoDoc = idxIng.tipoDoc !== undefined && ingresos[idxIng.tipoDoc] ? String(ingresos[idxIng.tipoDoc]).trim() : "13"; 
    filaSalida[2] = tipoDoc;
    
    let tipoDocNormalizado = tipoDoc.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
    filaSalida[9] = (tipoDocNormalizado === "CEDULA DE CIUDADANIA") ? "1" : "2"; 
    
    filaSalida[3] = rowHR[6]; 
    filaSalida[4] = rowHR[4]; 
    
    if (rowHR[25] instanceof Date) filaSalida[5] = Utilities.formatDate(rowHR[25], zonaHoraria, "yyyy-MM-dd");
    else filaSalida[5] = rowHR[25]; 

    filaSalida[6] = (String(rowHR[8]).toUpperCase() === "F") ? "1" : "2"; 
    filaSalida[7] = "0"; 
    filaSalida[8] = "0"; 
    
    filaSalida[10] = rowHR[27]; 
    filaSalida[11] = rowHR[17]; 
    filaSalida[12] = rowHR[46]; 
    
    if (rowHR[11] instanceof Date) filaSalida[13] = Utilities.formatDate(rowHR[11], zonaHoraria, "yyyy-MM-dd");
    else filaSalida[13] = rowHR[11]; 
    
    filaSalida[17] = "1"; filaSalida[18] = "2"; filaSalida[19] = "1"; filaSalida[20] = "001"; 
    
    filaSalida[23] = dicClasificador1[cargo] || "0";             
    filaSalida[24] = dicClasificador2[centroCosto] || "0";       
    filaSalida[25] = idxIng.clasif3 !== undefined ? String(ingresos[idxIng.clasif3] || "").trim() : ""; 
    filaSalida[26] = idxIng.prod !== undefined ? String(ingresos[idxIng.prod] || "").trim() : "0";    
    filaSalida[27] = idxIng.att !== undefined ? String(ingresos[idxIng.att] || "").trim() : "0";     
    filaSalida[28] = idxIng.site !== undefined ? String(ingresos[idxIng.site] || "").trim() : "0";     
    filaSalida[29] = dicClasificador7[centroCosto] || "";        
    
    filaSalida[30] = dicCodCargo[cargo] || "";
    
    filaSalida[32] = "306"; 
    filaSalida[33] = "0.522"; 
    
    filaSalida[38] = "1"; filaSalida[40] = "15"; filaSalida[63] = "1"; filaSalida[67] = "1"; filaSalida[70] = "1"; 
    
    filaSalida[21] = dicSucursales[centroCosto] || ""; 
    filaSalida[22] = centroCosto; 
    filaSalida[31] = tipoContrato;
    
    filaSalida[34] = dicFondosPension[String(fondos.pension).trim()] || fondos.pension; 
    filaSalida[35] = dicFondosSalud[String(fondos.eps).trim()] || fondos.eps;           
    filaSalida[37] = dicFondosCesantias[String(fondos.cesantias).trim()] || fondos.cesantias; 
    
    let ciudadHR = String(rowHR[idxHR.ciudad] || "").trim().toUpperCase();
    filaSalida[36] = dicFondoCCF[ciudadHR] || "";
    
    filaSalida[39] = salario;     
    
    if (tipoContrato === "10" || tipoContrato === "11") {
      filaSalida[43] = "52";
      filaSalida[69] = "52";
    } else {
      filaSalida[43] = dicGastos[centroCosto] || "";
      filaSalida[69] = dicGastos[centroCosto] || "";
    }
    
    filaSalida[15] = dicBancos[bancoTexto] || bancoTexto; 
    
    let infoCuenta = idxIng.cuenta !== undefined ? String(ingresos[idxIng.cuenta] || "") : "";
    filaSalida[16] = infoCuenta; 
    filaSalida[14] = (infoCuenta !== "") ? "1" : "4"; 
    
    filaSalida[46] = (tipoContrato === "10") ? "12" : (tipoContrato === "11" ? "1" : "0"); 
    
    let ciudExp = idxIng.ciudadExp !== undefined ? String(ingresos[idxIng.ciudadExp] || "") : "";
    let ciudNac = idxIng.ciudadNac !== undefined ? String(ingresos[idxIng.ciudadNac] || "") : "";
    
    filaSalida[47] = extraerPais(ciudExp);
    filaSalida[48] = ciudExp; 
    filaSalida[49] = extraerPais(ciudNac); 
    filaSalida[50] = ciudNac; 
    
    filaSalida[51] = "169"; filaSalida[52] = rowHR[idxHR.ciudad]; filaSalida[53] = "169"; filaSalida[54] = rowHR[idxHR.ciudad]; 
    filaSalida[55] = "169"; filaSalida[56] = rowHR[idxHR.ciudad]; 
    
    filaSalida[59] = (tipoContrato === "11") ? "19" : "01"; 
    filaSalida[66] = horasMesCalculado; 
    filaSalida[76] = valorHoraCalc;     
    filaSalida[77] = (idPagoElectronico === cedula) ? "" : idPagoElectronico; 
    
    filaSalida[84] = rowHR[7]; filaSalida[85] = rowHR[5]; 
    
    let factorRhVal = idxIng.rh !== undefined ? String(ingresos[idxIng.rh] || "").trim() : "";
    filaSalida[88] = factorRhVal ? factorRhVal.slice(0, -1) : ""; 
    filaSalida[89] = factorRhVal ? factorRhVal.slice(-1) : "";   
    filaSalida[91] = ""; 
    
    let fechaFinCelda = idxIng.fechaFin !== undefined ? ingresos[idxIng.fechaFin] : "";
    let fechaFinFormatted = "";
    if (fechaFinCelda instanceof Date) {
      fechaFinFormatted = Utilities.formatDate(fechaFinCelda, zonaHoraria, "yyyy-MM-dd");
    } else if (fechaFinCelda) {
      fechaFinFormatted = String(fechaFinCelda).trim();
    }
    filaSalida[94] = fechaFinFormatted;
    
    filaSalida[95] = dicCentroTrabajo[centroCosto] || "1"; 
    filaSalida[96] = dicCodigoArea[centroCosto] || "0"; 
    
    filaSalida[116] = rowHR[idxHR.legal] ? String(rowHR[idxHR.legal]).trim() : ""; 

    resultadoFinal.push(filaSalida);
  }
  
  if (resultadoFinal.length > 1) {
    sheetDatos.getRange(1, 1, resultadoFinal.length, 118).setValues(resultadoFinal);
    ui.alert(`✅ ¡ÉXITO!\n\nSe generó el cargue de Novasoft con ${resultadoFinal.length - 1} empleados.\nCédula de ciudadanía validada estrictamente, codCargo aplicado y Cargo(Title) extraído directamente de HR.`);
  } else {
    ui.alert(`❌ No se encontraron empleados para la fecha: ${fechaSeleccionada}`);
  }
}