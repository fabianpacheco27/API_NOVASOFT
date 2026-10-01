/**
 * GENERADOR DE CARGUE NOVASOFT V3
 * VERSIÓN MODULAR: Diccionarios extraídos a Google Sheets.
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
  
  // ---> ID DE TU NUEVA MATRIZ DE DICCIONARIOS <---
  const idDiccionarios = "1UtagOI9XtFoSdF1mC8uNvYVE6QxyktCkg0GTufrgt_U"; 
  const appDiccionarios = SpreadsheetApp.openById(idDiccionarios);
  
  let sheetDatos = ss.getSheetByName('DATOS');
  if (!sheetDatos) sheetDatos = ss.insertSheet('DATOS');
  sheetDatos.clear();
  
  // ==========================================
  // 2. CARGA DINÁMICA DE DICCIONARIOS
  // ==========================================
  // Esta función lee cualquier pestaña de tu matriz y la convierte en un diccionario instantáneo
  function cargarDiccionario(nombreHoja) {
    let sheet = appDiccionarios.getSheetByName(nombreHoja);
    if (!sheet) return {}; 
    let data = sheet.getDataRange().getValues();
    let dict = {};
    // Empezamos en i=1 para saltarnos el encabezado (Llave / Valor)
    for (let i = 1; i < data.length; i++) {
      let key = String(data[i][0]).trim();
      let val = String(data[i][1]).trim();
      if (key !== "") dict[key] = val;
    }
    return dict;
  }

  const dicBancos = cargarDiccionario("dicBancos");
  const dicSucursales = cargarDiccionario("dicSucursales");
  const dicGastos = cargarDiccionario("dicGastos");
  const dicFondosPension = cargarDiccionario("dicFondosPension");
  const dicFondosSalud = cargarDiccionario("dicFondosSalud");
  const dicFondosCesantias = cargarDiccionario("dicFondosCesantias");
  const dicClasificador1 = cargarDiccionario("dicClasificador1");
  const dicCodCargo = cargarDiccionario("dicCodCargo");
  const dicClasificador2 = cargarDiccionario("dicClasificador2");
  const dicClasificador7 = cargarDiccionario("dicClasificador7");
  const dicCentroTrabajo = cargarDiccionario("dicCentroTrabajo");
  const dicCodigoArea = cargarDiccionario("dicCodigoArea");
  const dicFondoCCF = cargarDiccionario("dicFondoCCF");
  const dicPaises = cargarDiccionario("dicPaises");

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
    ui.alert(`✅ ¡ÉXITO!\n\nCargue generado con ${resultadoFinal.length - 1} empleados.\nLos diccionarios se cargaron dinámicamente desde tu Google Sheet externo.`);
  } else {
    ui.alert(`❌ No se encontraron empleados para la fecha: ${fechaSeleccionada}`);
  }
}