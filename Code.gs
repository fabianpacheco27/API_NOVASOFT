/**
 * GENERADOR DE CARGUE NOVASOFT V2 
 * (CON MEGA-DICCIONARIOS VERTICALES Y MARCAS DE "PENDIENTE")
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
  // 2. DICCIONARIOS
  // ==========================================
  
  const dicBancos = {
    "BANCO DE LA REPÚBLICA": "00",
    "BANCO DE BOGOTÁ": "01",
    "BANCO POPULAR": "02",
    "ITAÚ CORPBANCA COLOMBIA S.A.": "06",
    "BANCOLOMBIA S.A.": "07",
    "CITIBANK COLOMBIA": "09",
    "HSBC": "10",
    "GNB SUDAMERIS S.A.": "12",
    "BBVA COLOMBIA": "13",
    "BANCO DE CREDITO": "14",
    "SCOTIABANK COLPATRIA": "19",
    "BANCO DE OCCIDENTE": "23",
    "BCSC S.A.": "30",
    "BANCO CAJA SOCIAL - BCSC S.A.": "32",
    "INTERNACIONAL CIA FINANCIAMIEN": "33",
    "Banco Produban": "36",
    "BANCO AGRARIO DE COLOMBIA S.A.": "40",
    "BANCO DAVIVIENDA S.A.": "51",
    "BANCO AV VILLAS": "52",
    "BANCO W S.A.": "53",
    "BANCO COLMENA": "57",
    "BANCO PROCREDIT": "58",
    "BANCAMIA": "59",
    "BANCO PICHINCHA S.A.": "60",
    "BANCOOMEVA": "61",
    "CMR FALABELLA S.A.": "62",
    "BANCO FINANDINA S.A.": "63",
    "BANCO MULTIBANK S.A": "64",
    "BANCO SANTANDER COLOMBIA S.A.": "65",
    "BANCO COOPERATIVO COOPCENTRAL": "66",
    "BANCO COMPARTIR S.A": "67",
    "BANCO SERFINANZA S.A": "69",
    "LULO BANK S.A.": "70",
    "NEQUI": "74",
    "DAVIPLATA": "75",
    "FINANCIERA JURISCOOP": "76",
    "COOP FINANCIERA DE ANTIOQUIA": "77",
    "COOTRAFA COOP FINANCIERA": "78",
    "CONFIAR COOPERATIVA FINANCIERA": "79",
    "COLTEFINANCIERA S.A": "80",
    "EC PACIFIC": "81",
    "Banco Correval": "90",
    "Credicorp Capital": "A1",
    "RAPPIPAY": "15",
    "Ualá": "16",
    "NU BANK": "82",
    "NU": "82",
    "Dale": "17"
  };

  const dicSucursales = {
    "Accounting and Payroll Admin": "35",
    "AIRBNB Admin": "35",
    "AIRBNB Agent Backoffice": "35",
    "AIRBNB Agent Voice": "35",
    "AIRBNB Agente Email": "35",
    "BOOKING Admin": "35",
    "BOOKING Agente Backoffice": "35",
    "BOOKING Agente Email": "35",
    "BOOKING Agente Voice": "35",
    "Campus Operations Admin": "35",
    "Campus Recruiting Operations Admin": "35",
    "Corporate Recruiting Admin": "35",
    "Customer Experience Admin": "35",
    "EX Admin": "35",
    "Facilities Admin": "35",
    "Global Operations Admin": "35",
    "GNC Admin": "35",
    "GNC Agente Backoffice": "35",
    "GNC Agente Email": "35",
    "GNC Agente Voice": "35",
    "Human Resources Admin": "35",
    "IT Operations Admin": "35",
    "LEVIS Admin": "35",
    "LEVIS Agente Backoffice": "35",
    "LEVIS Agente Email": "35",
    "LEVIS Agente Voice": "35",
    "LOREAL Admin": "35",
    "LOREAL Agente Backoffice": "35",
    "LOREAL Agente Email": "35",
    "LOREAL Agente Voice": "35",
    "Marketing Admin": "35",
    "Operations Leadership Admin": "35",
    "OPTAVIA Admin": "35",
    "OPTAVIA Agente Backoffice": "35",
    "OPTAVIA Agente Email": "35",
    "OPTAVIA Agente Voice": "35",
    "SEPHORA Admin": "20",
    "SEPHORA Agente Backoffice": "20",
    "SEPHORA Agente Email": "20",
    "SEPHORA Agente Voice": "20",
    "SPARK Agente Backoffice": "35",
    "SPARK Agente Chat": "35",
    "SPARK Chat Admin": "35",
    "SPARK Agente Email": "35",
    "SPARK Agente Voice": "35",
    "SPARK Email Admin": "35",
    "SPARK Voice Admin": "35",
    "TCP Admin": "35",
    "TCP Agente Backoffice": "35",
    "TCP Agente Email": "35",
    "TCP Agente Voice": "35",
    "Training Admin": "35",
    "VROOM Admin": "35",
    "VROOM Agente Backoffice": "35",
    "VROOM Agente Email": "35",
    "VROOM Agente Voice": "35",
    "WEBER Admin": "35",
    "WEBER Agente Backoffice": "35",
    "WEBER Agente Email": "35",
    "WEBER Agente Voice": "35",
    "MEJURI Admin": "35",
    "MEJURI Agente Backoffice": "35",
    "RED ROBIN Admin": "35",
    "RED ROBIN Agente Backoffice": "35",
    "Marigold Admin": "35",
    "Marigold Agente Voice": "35",
    "Noom Admin": "35",
    "Noom Agente Voice": "35",
    "CarGurus Admin": "35",
    "CarGurus Agente Voice": "35",
    "NYT Admin": "35",
    "NYT Agente Voice": "35",
    "Newell Admin": "35",
    "Newell Agente Voice": "35",
    "StubHub Admin": "35",
    "StubHub Agente Voice": "35",
    "HCA Servicing Admin": "35",
    "HCA Servicing Agente Voice": "35",
    "Workforce": "35",
    "EnFin Agente Voice": "35",
    "EnFin Admin": "35",
    "Campus Global": "35",
    "Pandora Admin": "35",
    "Pandora Agente Voice": "35",
    "PetSmart Admin": "35",
    "PetSmart Agente Voice": "35",
    "NCAA Admin": "35",
    "NCAA Agente Voice": "35",
    "MLB Admin": "35",
    "MLB Agente Voice": "35"
  };

  const dicGastos = {
    "Accounting and Payroll Admin": "52",
    "AIRBNB Admin": "51",
    "AIRBNB Agent Backoffice": "61",
    "AIRBNB Agent Voice": "61",
    "AIRBNB Agente Email": "61",
    "BOOKING Admin": "51",
    "BOOKING Agente Backoffice": "61",
    "BOOKING Agente Email": "61",
    "BOOKING Agente Voice": "61",
    "Campus Operations Admin": "52",
    "Campus Recruiting Operations Admin": "52",
    "Corporate Recruiting Admin": "52",
    "Customer Experience Admin": "51",
    "EX Admin": "52",
    "Facilities Admin": "52",
    "Global Operations Admin": "52",
    "GNC Admin": "51",
    "GNC Agente Backoffice": "61",
    "GNC Agente Email": "61",
    "GNC Agente Voice": "61",
    "Human Resources Admin": "52",
    "IT Operations Admin": "52",
    "LEVIS Admin": "51",
    "LEVIS Agente Backoffice": "61",
    "LEVIS Agente Email": "61",
    "LEVIS Agente Voice": "61",
    "LOREAL Admin": "51",
    "LOREAL Agente Backoffice": "61",
    "LOREAL Agente Email": "61",
    "LOREAL Agente Voice": "61",
    "Marketing Admin": "52",
    "Operations Leadership Admin": "52",
    "OPTAVIA Admin": "51",
    "OPTAVIA Agente Backoffice": "61",
    "OPTAVIA Agente Email": "61",
    "OPTAVIA Agente Voice": "61",
    "SEPHORA Admin": "51",
    "SEPHORA Agente Backoffice": "61",
    "SEPHORA Agente Email": "61",
    "SEPHORA Agente Voice": "61",
    "SPARK Agente Backoffice": "61",
    "SPARK Agente Chat": "61",
    "SPARK Chat Admin": "51",
    "SPARK Agente Email": "61",
    "SPARK Agente Voice": "61",
    "SPARK Email Admin": "51",
    "SPARK Voice Admin": "51",
    "TCP Admin": "51",
    "TCP Agente Backoffice": "61",
    "TCP Agente Email": "61",
    "TCP Agente Voice": "61",
    "Training Admin": "52",
    "VROOM Admin": "51",
    "VROOM Agente Backoffice": "61",
    "VROOM Agente Email": "61",
    "VROOM Agente Voice": "61",
    "WEBER Admin": "51",
    "WEBER Agente Backoffice": "61",
    "WEBER Agente Email": "61",
    "WEBER Agente Voice": "61",
    "MEJURI Admin": "51",
    "MEJURI Agente Backoffice": "61",
    "RED ROBIN Admin": "51",
    "RED ROBIN Agente Backoffice": "61",
    "Marigold Admin": "51",
    "Marigold Agente Voice": "61",
    "Noom Admin": "51",
    "Noom Agente Voice": "61",
    "CarGurus Admin": "51",
    "CarGurus Agente Voice": "61",
    "NYT Admin": "51",
    "NYT Agente Voice": "61",
    "Newell Admin": "51",
    "Newell Agente Voice": "61",
    "StubHub Admin": "51",
    "StubHub Agente Voice": "61",
    "HCA Servicing Admin": "51",
    "HCA Servicing Agente Voice": "61",
    "Workforce": "51",
    "EnFin Agente Voice": "61",
    "EnFin Admin": "51",
    "Campus Global": "51",
    "Pandora Admin": "51",
    "Pandora Agente Voice": "61",
    "PetSmart Admin": "51",
    "PetSmart Agente Voice": "61",
    "NCAA Admin": "51",
    "NCAA Agente Voice": "61",
    "MLB Admin": "51",
    "MLB Agente Voice": "61"
  };

  const dicFondosPension = {
    "ADMINISTRADORA COLOMBIANA DE PENSIONES COLPENSIONES": "108",
    "CAJA DE AUXILIOS Y DE PRESTACIONES DE ACDAC": "105",
    "COLFONDOS": "104",
    "FONDO DE PREVISIÓN SOCIAL DEL CONGRESO": "106",
    "NO APLICA (APRENDICES)": "0",
    "NO APLICA PENSION": "199",
    "SKANDIA": "102",
    "PENSIONES DE ANTIOQUIA": "107",
    "PORVENIR": "101",
    "PROTECCIÓN": "100",
    "OLD MUTUAL FONDO DE PENSIONES OBLIGATORIAS": "102"
  };

  const dicFondosSalud = {
    "ALIANSALUD EPS": "200",
    "ASMET SALUD EPS": "224",
    "CAJA DE PREVISIÓN SOCIAL DE CASANARE - CAPRESOCA E.P.S.": "231",
    "CAJA DE PREVISIÓN SOCIAL DE COMUNICACIONES - CAPRECOM": "230",
    "PROTEGER EPS S.A.S": "225",
    "CAPITAL SALUD": "218",
    "COMFACUNDI": "220",
    "COMFAMILIAR HUILA": "227",
    "COMFENALCO VALLE EPS": "206",
    "COMPENSAR ENTIDAD PROMOTORA DE SALUD": "204",
    "COMFAORIENTE EPS-S": "234",
    "COOSALUD EPS": "219",
    "E.P.S SANITAS": "203",
    "EMPRESA COOPERATIVA SOLIDARIA DE SALUD “ECOOPSOS”": "223",
    "EMPRESAS PÚBLICAS DE MEDELLÍN DEPARTAMENTO MÉDICO": "216",
    "EMSSANAR": "221",
    "EPS SURA": "205",
    "FAMISANAR": "209",
    "FONDO DE PASIVO SOCIAL DE FERROCARRILES": "217",
    "FONDO DE SOLIDARIDAD Y GARANTÍA FOSYGA": "215",
    "Fundación Salud Mía EPS": "229",
    "MUTUAL SER E.S.S.": "222",
    "NO APLICA": "299",
    "NUEVA EPS": "214",
    "SALUD TOTAL S.A.": "201",
    "SALUDVIDA S.A EPS": "213",
    "SAVIA SALUD EPS": "228",
    "SERVICIO OCCIDENTAL DE SALUD S.A. S.O.S.": "210",
    "SALUD BOLÍVAR EPS S.A.S": "233",
    "EPS FAMILIAR DE COLOMBIA S.A.S.": "236"
  };

  const dicFondosCesantias = {
    "CESANTIAS COLFONDOS": "503",
    "CESANTIAS PORVENIR": "502",
    "CESANTIAS PROTECCIÓN": "501",
    "CESANTIAS SKANDIA": "504",
    "FONDO NACIONAL DEL AHORRO - FNA": "500",
    "NO APLICA": "599"
  };
  
  // 3. LEEMOS DATA
  const dataHR = sheetHR.getDataRange().getValues();
  const dataIngresosDisplay = sheetIngresos.getDataRange().getDisplayValues();
  const dataIngresosValores = sheetIngresos.getDataRange().getValues(); 
  const dataStaff = sheetStaff.getDataRange().getValues();
  const dataAgentes = sheetAgentes.getDataRange().getValues();
  
  let ingresosPorCedula = {};
  let fechaRealPorCedula = {}; 
  for (let i = 1; i < dataIngresosDisplay.length; i++) {
    let cedula = String(dataIngresosDisplay[i][10]).trim(); 
    ingresosPorCedula[cedula] = dataIngresosDisplay[i];
    fechaRealPorCedula[cedula] = dataIngresosValores[i][0]; 
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
  
  // 4. ENCABEZADOS (118)
  const encabezados = [
    'codEmp', 'nroIdentificacion', 'tipoIdentificacion', 'primerApellido', 'primerNombre', 'fechaNacimiento', 'sexoEmp', 'claseLibreta', 'estadocivil', 'nacionalidad', 'direccionRes', 'email', 'celular', 'fechaIngreso', 'tipoPago', 'banco', 'cuentaBanco', 'tipoLIquida', 'regimenSal', 'claseSalario', 'compania', 'Sucursal', 'centroCosto', 'clasificador1', 'clasificador2', 'clasificador3', 'clasificador4', 'clasificador5', 'clasificador6', 'clasificador7', 'codCargo', 'tipoContrato', 'fondoRiesgos', 'porRiesgos', 'fondoPension', 'fondoSalud', 'fondoCCF', 'fondoCesant', 'metReten', 'salario', 'diasVacacionAno', 'sabadoHabil', 'promedioSalud', 'cuentaGasto', 'variable', 'porDias', 'modoLiquidacion', 'paisExpIdentif', 'ciudadExpIdent', 'paisNacimiento', 'ciudadNacimiento', 'paisRes', 'ciudadRes', 'paisLab', 'ciuLabor', 'paisContrato', 'ciuContrato', 'pensionado', 'pensxEmpresa', 'tipCotizante', 'subtipoCotizante', 'extranjero', 'resideExtranjero', 'ley1393', 'contratista', 'ley1450', 'horasMes', 'nroContrato', 'sucursalSS', 'cuentaGastoNiif', 'distribucionCCosto', 'distriCCostoNiif', 'deducibleSalud', 'deducibleVivienda', 'dependientes', 'nroPersoCargo', 'valorHora', 'codPagoElectronico', 'PagarDia31Vac', 'Localidad', 'CodBarrio', 'NroPasaporte', 'PaisEmisorPasaporte', 'codigoAlterno', 'segundoApellido', 'segundoNombre', 'numLibreta', 'distritoLibreta', 'grupoSang', 'factorRh', 'telefonoRes', 'barrio', 'porcReten', 'convenio', 'fechaFinContrato', 'codCentDeTrabajo', 'codigoArea', 'sueldoAntesFlex', 'porcenFlex', 'estatura', 'peso', 'emailAlterno', 'sucursalConvenio', 'ccostoConvenio', 'clasificador1Conv', 'clasificador2Conv', 'clasificador3Conv', 'clasificador4Conv', 'clasificador5Conv', 'clasificador6Conv', 'clasificador7Conv', 'clasificador8Conv', 'lider', 'FechaExpdocId', 'FechaExpPasaporte', 'ContratacionExterna', 'TipoTrabajo', 'IndTransicion'
  ];
  
  let resultadoFinal = [encabezados];
  const indicesCeros = [41, 42, 44, 45, 57, 58, 60, 61, 62, 64, 65, 68, 71, 72, 73, 74, 75, 78, 79, 80, 81, 82, 87, 92, 93, 97, 98, 99, 100, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 117];
  
  // 5. PROCESAMOS DATA
  for (let i = 1; i < dataHR.length; i++) {
    let rowHR = dataHR[i];
    let cedula = String(rowHR[23]).trim();
    if (!cedula || cedula === "") continue; 
    
    let fechaCelda = fechaRealPorCedula[cedula];
    if (!(fechaCelda instanceof Date)) continue; 
    let fechaFormateadaExcel = Utilities.formatDate(fechaCelda, zonaHoraria, "yyyy-MM-dd");
    
    if (fechaFormateadaExcel !== fechaSeleccionada) continue; 
    
    let ingresos = ingresosPorCedula[cedula] || [];
    let fondos = fondosPorCedula[cedula] || { eps: '', pension: '', cesantias: '' };
    
    let centroCosto = String(ingresos[4]).trim();
    let tipoContrato = ingresos[14] ? String(ingresos[14]).trim() : ""; 
    let cargo = ingresos[6] ? String(ingresos[6]) : ""; 
    let bancoTexto = ingresos[21] ? String(ingresos[21]).trim() : "";
    
    let horasMesCalculado = (tipoContrato === "10" || tipoContrato === "11" || cargo.toUpperCase() !== "AGENTE TELEFONICO") ? 210 : (rowHR[41] || 210);
    let salario = ingresos[16] ? Number(ingresos[16]) : 0; 
    let valorHoraCalc = horasMesCalculado > 0 ? Math.round(salario / horasMesCalculado) : 0;
    let idPagoElectronico = ingresos[23] ? String(ingresos[23]).trim() : ""; 
    
    let filaSalida = new Array(118).fill(""); 
    indicesCeros.forEach(idx => filaSalida[idx] = "0");
    
    filaSalida[0] = rowHR[0]; 
    filaSalida[1] = cedula;   
    filaSalida[2] = "13";     
    filaSalida[3] = rowHR[6]; 
    filaSalida[4] = rowHR[4]; 
    
    if (rowHR[25] instanceof Date) filaSalida[5] = Utilities.formatDate(rowHR[25], zonaHoraria, "yyyy-MM-dd");
    else filaSalida[5] = rowHR[25]; 

    filaSalida[6] = (String(rowHR[8]).toUpperCase() === "F") ? "1" : "2"; 
    filaSalida[7] = "0"; filaSalida[8] = "1"; filaSalida[9] = "1"; 
    filaSalida[10] = rowHR[27]; 
    filaSalida[11] = rowHR[17]; 
    filaSalida[12] = rowHR[46]; 
    
    if (rowHR[11] instanceof Date) filaSalida[13] = Utilities.formatDate(rowHR[11], zonaHoraria, "yyyy-MM-dd");
    else filaSalida[13] = rowHR[11]; 
    
    filaSalida[17] = "1"; filaSalida[18] = "2"; filaSalida[19] = "1"; filaSalida[20] = "001"; 
    
    // --- [PENDIENTE] - CLASIFICADORES ---
    filaSalida[23] = "PENDIENTE"; // clasificador1
    filaSalida[24] = "PENDIENTE"; // clasificador2 (HR)
    filaSalida[25] = "PENDIENTE"; // clasificador3 (Ingresos TYPE)
    filaSalida[26] = "PENDIENTE"; // clasificador4
    filaSalida[27] = "PENDIENTE"; // clasificador5
    filaSalida[28] = "PENDIENTE"; // clasificador6
    filaSalida[29] = "PENDIENTE"; // clasificador7 (Ingresos TITLE)
    
    filaSalida[38] = "1"; filaSalida[40] = "15"; filaSalida[63] = "1"; filaSalida[67] = "1"; filaSalida[70] = "1"; 
    
    filaSalida[21] = dicSucursales[centroCosto] || ""; 
    filaSalida[22] = ingresos[4]; 
    filaSalida[31] = tipoContrato;
    
    filaSalida[34] = dicFondosPension[String(fondos.pension).trim()] || fondos.pension; 
    filaSalida[35] = dicFondosSalud[String(fondos.eps).trim()] || fondos.eps;           
    filaSalida[37] = dicFondosCesantias[String(fondos.cesantias).trim()] || fondos.cesantias; 
    
    filaSalida[39] = salario;     
    
    if (tipoContrato === "10" || tipoContrato === "11") {
      filaSalida[43] = "52";
      filaSalida[69] = "52";
    } else {
      filaSalida[43] = dicGastos[centroCosto] || "";
      filaSalida[69] = dicGastos[centroCosto] || "";
    }
    
    filaSalida[15] = dicBancos[bancoTexto] || bancoTexto; 
    filaSalida[16] = ingresos[20]; 
    filaSalida[14] = (filaSalida[16] && filaSalida[16] !== "") ? "1" : "4"; 
    
    filaSalida[46] = (tipoContrato === "10") ? "12" : (tipoContrato === "11" ? "1" : "0"); 
    
    filaSalida[47] = "169"; filaSalida[48] = ingresos[11]; filaSalida[49] = "169"; filaSalida[50] = ingresos[12]; 
    filaSalida[51] = "169"; filaSalida[52] = rowHR[45]; filaSalida[53] = "169"; filaSalida[54] = rowHR[45]; 
    filaSalida[55] = "169"; filaSalida[56] = rowHR[45]; 
    
    filaSalida[59] = (tipoContrato === "11") ? "19" : "01"; 
    filaSalida[66] = horasMesCalculado; 
    filaSalida[76] = valorHoraCalc;     
    filaSalida[77] = (idPagoElectronico === cedula) ? "" : idPagoElectronico; 
    
    filaSalida[84] = rowHR[7]; filaSalida[85] = rowHR[5]; 
    filaSalida[88] = ingresos[7] ? String(ingresos[7]).trim().slice(0, -1) : ""; 
    filaSalida[89] = ingresos[7] ? String(ingresos[7]).trim().slice(-1) : "";    
    filaSalida[91] = ""; 
    
    // --- [PENDIENTE] - CENTRO DE TRABAJO (RIESGO) Y ÁREA ---
    filaSalida[95] = "PENDIENTE"; // codCentDeTrabajo
    filaSalida[96] = "PENDIENTE"; // codigoArea
    
    filaSalida[116] = rowHR[49] ? String(rowHR[49]).trim() : ""; 

    resultadoFinal.push(filaSalida);
  }
  
  if (resultadoFinal.length > 1) {
    sheetDatos.getRange(1, 1, resultadoFinal.length, 118).setValues(resultadoFinal);
    ui.alert(`¡ÉXITO! Se generó el cargue Novasoft. Registros: ${resultadoFinal.length - 1} \n\n⚠️ Recuerda revisar las columnas marcadas como 'PENDIENTE'.`);
  } else {
    ui.alert(`❌ No se encontraron empleados con esa fecha.`);
  }
}