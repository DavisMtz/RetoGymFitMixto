/**
 * BACKEND COMPLETO - GYM RETO 2026
 * Incluye App Web y Panel de Administración Lateral
 */

// -----------------------------------------------------------
// 1. FUNCIONES DE ADMINISTRACIÓN (MENÚ Y SIDEBAR)
// -----------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('ADMIN GYM')
      .addItem('💰 Registrar Pago', 'mostrarSidebarPagos')
      .addToUi();
}

function mostrarSidebarPagos() {
  var html = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle('Registrar Pago 💰')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function guardarPago(datos) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Pagos");
    
    // Si no existe la hoja, avisar para evitar error
    if (!hoja) return { success: false, message: "Error: No existe la hoja 'Pagos'." };
    
    // Guardar: Fecha (Timestamp), Usuario, Monto, Notas
    // La fecha se guarda con new Date() del momento de registro
    hoja.appendRow([new Date(), datos.usuario, datos.monto, datos.notas]);
    
    return { success: true, message: "Pago registrado exitosamente." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// -----------------------------------------------------------
// 2. FUNCIONES DE LA WEB APP (FRONTEND USUARIOS)
// -----------------------------------------------------------

// ==========================================
// 1. CONFIGURACIÓN INICIAL Y VISUALIZACIÓN
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('GYM - RETO 2026')
      // MEJORA: Viewport bloqueado para evitar zoom accidental en inputs
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      // MEJORA: Etiquetas para que se comporte como App en iOS (Full Screen)
      .addMetaTag('apple-mobile-web-app-capable', 'yes')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 2. GESTIÓN DE USUARIOS
// ==========================================

// Obtener lista de atletas activos
function obtenerUsuarios() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Usuarios");
  if (!hoja) return [];
  var datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
  return datos.filter(r => r[0] !== "" && r[1] === "Activo").map(r => r[0]);
}

// Verificar si el usuario ya registró actividad hoy
function verificarEstadoHoy(usuario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Registros");
  var datos = hoja.getDataRange().getValues();
  
  // Fecha actual en string YYYY-MM-DD (Zona Horaria México)
  var hoy = Utilities.formatDate(new Date(), "America/Mexico_City", "yyyy-MM-dd");
  
  // Buscar si existe un registro de este usuario con la fecha de hoy
  for (var i = 1; i < datos.length; i++) {
    var usuarioFila = datos[i][1];
    var fechaFilaRaw = new Date(datos[i][2]);
    var fechaFila = Utilities.formatDate(fechaFilaRaw, "America/Mexico_City", "yyyy-MM-dd");
    
    if (usuarioFila === usuario && fechaFila === hoy) {
      return { 
        yaRegistro: true, 
        mensaje: "¡Ya registraste tu actividad de hoy!",
        tipo: datos[i][3] // Devolvemos qué hizo (Gym, Outdoor, etc)
      };
    }
  }
  
  return { yaRegistro: false };
}

// ==========================================
// 3. NUEVA FUNCIÓN: RACHAS (GAMIFICACIÓN)
// ==========================================

function obtenerRachaUsuario(usuario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Registros");
  var datos = hoja.getDataRange().getValues();
  
  // Filtrar fechas únicas donde el usuario cumplió
  var fechasCumplidas = [];
  var zona = "America/Mexico_City";
  
  for (var i = 1; i < datos.length; i++) {
    // Verificamos Usuario Y que el estatus sea válido (CUMPLE o JUSTIFICADO)
    if (datos[i][1] === usuario && 
       (datos[i][7] === "CUMPLE" || datos[i][7] === "JUSTIFICADO")) {
      
      var f = Utilities.formatDate(new Date(datos[i][2]), zona, "yyyy-MM-dd");
      // Evitamos duplicados (si entrenó 2 veces el mismo día)
      if (fechasCumplidas.indexOf(f) === -1) {
        fechasCumplidas.push(f);
      }
    }
  }
  
  // Si no hay registros, racha es 0
  if (fechasCumplidas.length === 0) return 0;

  // Ordenar fechas de la más reciente a la más antigua
  fechasCumplidas.sort(function(a, b) {
    return new Date(b) - new Date(a);
  });

  var hoy = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");
  
  // Calcular AYER
  var ayerDate = new Date();
  ayerDate.setDate(ayerDate.getDate() - 1);
  var ayer = Utilities.formatDate(ayerDate, zona, "yyyy-MM-dd");

  // REGLA: Para mantener la racha viva, el último registro debe ser HOY o AYER.
  // Si el último registro fue anteayer, la racha ya se perdió.
  if (fechasCumplidas[0] !== hoy && fechasCumplidas[0] !== ayer) {
    return 0;
  }

  // Contar días consecutivos hacia atrás
  var racha = 1; // Empezamos con 1 porque ya validamos el último día
  
  for (var i = 0; i < fechasCumplidas.length - 1; i++) {
    var fechaReciente = new Date(fechasCumplidas[i]);
    var fechaAnterior = new Date(fechasCumplidas[i+1]);
    
    // Diferencia en milisegundos
    var diffTime = Math.abs(fechaReciente - fechaAnterior);
    // Convertir a días (redondeando hacia arriba para asegurar)
    var diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24)); 
    
    if (diffDays === 1) {
      racha++;
    } else {
      // Si hay un salto de más de 1 día, se rompe la cadena
      break; 
    }
  }
  
  return racha;
}

// ==========================================
// 4. HISTORIAL Y GUARDADO
// ==========================================

function obtenerHistorialUsuario(usuario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Registros");
  var datos = hoja.getDataRange().getValues();
  var historial = [];

  for (var i = datos.length - 1; i >= 1; i--) {
    if (datos[i][1] === usuario) {
      var fechaRaw = new Date(datos[i][2]);
      var fechaStr = Utilities.formatDate(fechaRaw, "America/Mexico_City", "dd/MM");
      
      historial.push({
        fecha: fechaStr,
        tipo: datos[i][3],
        estatus: datos[i][7]
      });
      if (historial.length >= 5) break;
    }
  }
  return historial;
}

function obtenerSemanaUsuario(usuario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Registros");
  var datos = hoja.getDataRange().getValues();
  var zona = "America/Mexico_City";
  var hoy = new Date();
  var hoyStr = Utilities.formatDate(hoy, zona, "yyyy-MM-dd");
  var hp = hoyStr.split('-');
  var hoyObj = new Date(hp[0], hp[1]-1, hp[2], 12, 0, 0);
  var diaNorm = hoyObj.getDay() === 0 ? 7 : hoyObj.getDay();
  var lunes = new Date(hoyObj); lunes.setDate(hoyObj.getDate() - (diaNorm - 1));
  var dias = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];
  var semana = [];
  for (var d = 0; d < 7; d++) {
    var f = new Date(lunes); f.setDate(lunes.getDate() + d);
    semana.push({ dia: dias[d], fecha: Utilities.formatDate(f, zona, "yyyy-MM-dd"), estatus: 'sin registro' });
  }
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][1] !== usuario) continue;
    var fechaReg = Utilities.formatDate(new Date(datos[i][2]), zona, "yyyy-MM-dd");
    for (var d = 0; d < 7; d++) {
      if (semana[d].fecha === fechaReg) { semana[d].estatus = datos[i][7]; break; }
    }
  }
  return semana;
}

function guardarRegistro(datos) {
  try {
    // 1. Doble validación servidor para evitar trampas/doble clic
    var estado = verificarEstadoHoy(datos.usuario);
    if (estado.yaRegistro) {
      return { success: false, message: "Ya existe un registro tuyo el día de hoy." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Registros");
    
    // Fecha normalizada
    var ahora = new Date();
    var fechaString = Utilities.formatDate(ahora, "America/Mexico_City", "yyyy-MM-dd");
    var partesFecha = fechaString.split('-');
    // Guardamos a mediodía para evitar problemas de zona horaria al leer
    var fechaGuardar = new Date(partesFecha[0], partesFecha[1] - 1, partesFecha[2], 12, 0, 0);
    
    // Cálculos
    var minTotales = 0;
    if (datos.tiempo) {
      var partes = datos.tiempo.split(':');
      minTotales = (parseInt(partes[0]) * 60) + parseInt(partes[1]);
    }
    var cal = parseInt(datos.calorias) || 0;
    var tipo = datos.tipo;
    var cumple = "Pendiente";

    // Reglas Lógicas
    if (tipo === "Gimnasio") {
      cumple = (minTotales >= 75 || cal >= 500) ? "CUMPLE" : "NO CUMPLE";
    } else if (tipo === "Fuera del Gym") {
      cumple = ((minTotales >= 90 && cal >= 550) || (cal >= 700 && minTotales <= 120)) ? "CUMPLE" : "NO CUMPLE";
    } else if (tipo === "Vacaciones" || tipo === "Incapacidad") {
      cumple = "JUSTIFICADO";
      minTotales = 0; cal = 0;
    } else { cumple = "N/A"; }

    hoja.appendRow([new Date(), datos.usuario, fechaGuardar, tipo, minTotales, cal, datos.evidencia, cumple, datos.notas]);
    
    return { success: true, message: "¡Entrenamiento registrado con éxito!" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// 5. RANKING Y BOTE
// ==========================================

function obtenerRanking() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaReg = ss.getSheetByName("Registros");
  var datos = hojaReg.getDataRange().getValues();

  // 1. CALCULAR RANGO DE FECHAS DE LA SEMANA ACTUAL (LUNES - DOMINGO)
  var hoy = new Date();
  var zona = "America/Mexico_City";
  
  var hoyStr = Utilities.formatDate(hoy, zona, "yyyy-MM-dd");
  var hoyPartes = hoyStr.split('-');
  var hoyObj = new Date(hoyPartes[0], hoyPartes[1]-1, hoyPartes[2], 12, 0, 0);
  
  var diaSemana = hoyObj.getDay(); 
  var diaNormalizado = (diaSemana === 0) ? 7 : diaSemana;
  
  var lunesSemana = new Date(hoyObj);
  lunesSemana.setDate(hoyObj.getDate() - (diaNormalizado - 1));
  lunesSemana.setHours(0,0,0,0);
  
  var domingoSemana = new Date(lunesSemana);
  domingoSemana.setDate(lunesSemana.getDate() + 6);
  domingoSemana.setHours(23,59,59,999);

  var conteoSemanal = {};
  
  // 2. PROCESAR REGISTROS
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var nombre = fila[1];
    var fechaReg = new Date(fila[2]); 
    
    if (fechaReg >= lunesSemana && fechaReg <= domingoSemana) {
      var estatus = fila[7]; 
      var kcal = Number(fila[5]) || 0; 
      var tipo = fila[3]; 

      if (!conteoSemanal[nombre]) conteoSemanal[nombre] = { dias: 0, calorias: 0 };
      
      if (estatus === "CUMPLE" || estatus === "JUSTIFICADO") {
        conteoSemanal[nombre].dias += 1;
      }
      if (tipo === "Gimnasio" || tipo === "Fuera del Gym") {
        conteoSemanal[nombre].calorias += kcal;
      }
    }
  }

  // 3. CALCULAR BOTE
  var boteTotal = 0;
  var hojaPagos = ss.getSheetByName("Pagos");
  
  if (hojaPagos && hojaPagos.getLastRow() > 1) {
    var datosPagos = hojaPagos.getRange(2, 1, hojaPagos.getLastRow()-1, 5).getValues(); 
    for (var p = 0; p < datosPagos.length; p++) {
      var monto = parseFloat(datosPagos[p][2]); 
      if (!isNaN(monto)) {
        boteTotal += monto;
      }
    }
  }

  // 4. ORDENAR Y RETORNAR
  var rankingFinal = [];
  for (var key in conteoSemanal) {
    rankingFinal.push({
      nombre: key,
      dias: conteoSemanal[key].dias,
      calorias: conteoSemanal[key].calorias
    });
  }
  
  rankingFinal.sort(function(a, b) {
    if (b.dias !== a.dias) return b.dias - a.dias;
    return b.calorias - a.calorias;
  });

  return {
    ranking: rankingFinal,
    bote: boteTotal
  };
}

function obtenerRankingMensual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaReg = ss.getSheetByName("Registros");
  var datos = hojaReg.getDataRange().getValues();
  var zona = "America/Mexico_City";
  var hoyStr = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");
  var partes = hoyStr.split('-');
  var anio = parseInt(partes[0]), mes = parseInt(partes[1]) - 1;
  var primerDia = new Date(anio, mes, 1, 0, 0, 0, 0);
  var ultimoDia = new Date(anio, mes + 1, 0, 23, 59, 59, 999);
  var conteo = {};
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i], nombre = fila[1], fechaReg = new Date(fila[2]);
    if (fechaReg >= primerDia && fechaReg <= ultimoDia) {
      if (!conteo[nombre]) conteo[nombre] = { dias: 0, calorias: 0 };
      if (fila[7] === "CUMPLE" || fila[7] === "JUSTIFICADO") conteo[nombre].dias++;
      if (fila[3] === "Gimnasio" || fila[3] === "Fuera del Gym") conteo[nombre].calorias += Number(fila[5]) || 0;
    }
  }
  var boteTotal = 0;
  var hojaPagos = ss.getSheetByName("Pagos");
  if (hojaPagos && hojaPagos.getLastRow() > 1) {
    var dp = hojaPagos.getRange(2, 1, hojaPagos.getLastRow()-1, 5).getValues();
    for (var p = 0; p < dp.length; p++) { var m = parseFloat(dp[p][2]); if (!isNaN(m)) boteTotal += m; }
  }
  var rankingFinal = [];
  for (var key in conteo) rankingFinal.push({ nombre: key, dias: conteo[key].dias, calorias: conteo[key].calorias });
  rankingFinal.sort(function(a,b) { return b.dias !== a.dias ? b.dias-a.dias : b.calorias-a.calorias; });
  return { ranking: rankingFinal, bote: boteTotal };
}

function obtenerRankingAnterior() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaReg = ss.getSheetByName("Registros");
  var datos = hojaReg.getDataRange().getValues();
  var zona = "America/Mexico_City";
  var hoyStr = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd");
  var hp = hoyStr.split('-');
  var hoyObj = new Date(hp[0], hp[1]-1, hp[2], 12, 0, 0);
  var diaNorm = hoyObj.getDay() === 0 ? 7 : hoyObj.getDay();
  var lunesEsta = new Date(hoyObj); lunesEsta.setDate(hoyObj.getDate() - (diaNorm - 1));
  var lunesAnt = new Date(lunesEsta); lunesAnt.setDate(lunesEsta.getDate() - 7);
  lunesAnt.setHours(0,0,0,0);
  var domingoAnt = new Date(lunesAnt); domingoAnt.setDate(lunesAnt.getDate() + 6); domingoAnt.setHours(23,59,59,999);
  var conteo = {};
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i], nombre = fila[1], fechaReg = new Date(fila[2]);
    if (fechaReg >= lunesAnt && fechaReg <= domingoAnt) {
      if (!conteo[nombre]) conteo[nombre] = { dias: 0, calorias: 0 };
      if (fila[7] === "CUMPLE" || fila[7] === "JUSTIFICADO") conteo[nombre].dias++;
      if (fila[3] === "Gimnasio" || fila[3] === "Fuera del Gym") conteo[nombre].calorias += Number(fila[5]) || 0;
    }
  }
  var rankingFinal = [];
  for (var key in conteo) rankingFinal.push({ nombre: key, dias: conteo[key].dias, calorias: conteo[key].calorias });
  rankingFinal.sort(function(a,b) { return b.dias !== a.dias ? b.dias-a.dias : b.calorias-a.calorias; });
  return { ranking: rankingFinal };
}

function obtenerActividadReciente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Registros");
  var datos = hoja.getDataRange().getValues();
  var ahora = new Date();
  var resultado = [];
  for (var i = datos.length - 1; i >= 1; i--) {
    var estatus = datos[i][7];
    if (estatus !== "CUMPLE" && estatus !== "JUSTIFICADO") continue;
    var diffMins = Math.round((ahora - new Date(datos[i][0])) / 60000);
    var hace = diffMins < 60 ? diffMins + ' min'
             : diffMins < 1440 ? Math.round(diffMins/60) + ' h'
             : Math.round(diffMins/1440) + ' d';
    resultado.push({
      nombre: datos[i][1].split(' ')[0],
      tipo: datos[i][3],
      calorias: Number(datos[i][5]) || 0,
      hace: hace
    });
    if (resultado.length >= 8) break;
  }
  return resultado;
}

// Función auxiliar: Número de semana ISO
function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
  return weekNo;
}
