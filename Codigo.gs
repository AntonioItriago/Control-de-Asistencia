// ==================== RECIBIR PETICIONES ====================
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (error) {
    data = e.parameter;
  }

  let respuesta = {};

  switch (data.accion) {
    case 'login':
      respuesta = procesarLogin(data.usuario, data.password);
      break;
    case 'registrar_asistencia':
      respuesta = registrarAsistencia(data);
      break;
    case 'crear_empleado':
      respuesta = crearEmpleado(data);
      break;
    case 'listar_empleados':
      respuesta = listarEmpleados();
      break;
    case 'eliminar_empleado':
      respuesta = eliminarEmpleado(data.usuario);
      break;
    case 'editar_empleado':
      respuesta = editarEmpleado(data);
      break;
    case 'listar_asistencias':
      respuesta = listarAsistencias();
      break;
    case 'cambiar_password':
      respuesta = cambiarPassword(data.usuario, data.passwordActual, data.passwordNueva);
      break;
    case 'guardar_empresa':
      respuesta = guardarEmpresa(data);
      break;
    case 'obtener_empresa':
      respuesta = obtenerEmpresa();
      break;
    default:
      respuesta = { exito: false, mensaje: "Acción no reconocida: " + data.accion };
  }

  return ContentService.createTextOutput(JSON.stringify(respuesta))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return doPost(e);
}

// ==================== LOGIN ====================
function procesarLogin(usuario, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  if (!sheet) return { exito: false, mensaje: "No existe hoja 'Usuarios'" };
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(usuario).trim() && 
        String(datos[i][1]).trim() === String(password).trim()) {
      return { 
        exito: true, 
        rol: String(datos[i][5]).trim().toLowerCase(),
        nombre: String(datos[i][2]).trim(),
        area: String(datos[i][3]).trim(),
        cargo: String(datos[i][4]).trim()
      };
    }
  }
  return { exito: false, mensaje: "Usuario o contraseña incorrectos" };
}

// ==================== REGISTRAR ASISTENCIA ====================
function registrarAsistencia(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Asistencias');
  if (!sheet) return { exito: false, mensaje: "No existe hoja 'Asistencias'" };
  
  const datos = sheet.getDataRange().getValues();
  const ahora = new Date();
  
  // Usar formato ISO para comparación (más confiable)
  const fechaHoyISO = Utilities.formatDate(ahora, 'America/Caracas', 'yyyy-MM-dd');
  const fechaHoyDisplay = Utilities.formatDate(ahora, 'America/Caracas', 'dd/MM/yyyy');
  const horaAhora = Utilities.formatDate(ahora, 'America/Caracas', 'HH:mm:ss');
  const ubicacion = data.lat.toFixed(6) + ', ' + data.lng.toFixed(6);
  
  Logger.log('=== DEBUG registrarAsistencia ===');
  Logger.log('Usuario: ' + data.usuario);
  Logger.log('Tipo: ' + data.tipo);
  Logger.log('Fecha hoy ISO: ' + fechaHoyISO);
  Logger.log('Fecha hoy Display: ' + fechaHoyDisplay);
  
  let filaRegistro = -1;
  
  for (let i = 1; i < datos.length; i++) {
    const usuarioFila = String(datos[i][1]).trim();
    const fechaFilaRaw = datos[i][3];
    
    // Convertir fecha de la fila a formato ISO para comparar
    let fechaFilaISO = '';
    if (fechaFilaRaw instanceof Date) {
      fechaFilaISO = Utilities.formatDate(fechaFilaRaw, 'America/Caracas', 'yyyy-MM-dd');
    } else {
      const fechaStr = String(fechaFilaRaw).trim();
      // Si es formato dd/MM/yyyy, convertir a ISO
      if (fechaStr.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        const partes = fechaStr.split('/');
        fechaFilaISO = partes[2] + '-' + partes[1] + '-' + partes[0];
      } else if (fechaStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
        fechaFilaISO = fechaStr;
      } else {
        fechaFilaISO = fechaStr;
      }
    }
    
    Logger.log('Fila ' + i + ': usuario=' + usuarioFila + ', fechaRaw=' + fechaFilaRaw + ', fechaISO=' + fechaFilaISO);
    
    if (usuarioFila === data.usuario && fechaFilaISO === fechaHoyISO) {
      filaRegistro = i + 1;
      Logger.log('✅ ENCONTRADO en fila ' + filaRegistro);
      break;
    }
  }
  
  Logger.log('Resultado filaRegistro: ' + filaRegistro);
  
  try {
    switch (data.tipo) {
      case 'entrada':
        if (filaRegistro !== -1) {
          Logger.log('❌ Ya existe registro hoy');
          return { exito: false, mensaje: "Ya registraste entrada hoy." };
        }
        let fotoEntradaUrl = '';
        if (data.foto && data.foto.length > 100) {
          fotoEntradaUrl = guardarFoto(data.foto, data.usuario + '_entrada_' + Date.now());
        }
        sheet.appendRow(['AS-' + ahora.getTime(), data.usuario, data.nombre || data.usuario, fechaHoyDisplay, horaAhora, ubicacion, fotoEntradaUrl, '', '', '', '', '', '', '', '', '', '', '']);
        Logger.log('✅ Entrada registrada');
        return { exito: true, mensaje: 'Entrada registrada a las ' + horaAhora };
      
      case 'comida-in':
        if (filaRegistro === -1) {
          Logger.log('❌ No hay registro de hoy');
          return { exito: false, mensaje: "Primero registra la Entrada." };
        }
        if (sheet.getRange(filaRegistro, 8).getValue()) {
          return { exito: false, mensaje: "Ya registraste Comida-In." };
        }
        sheet.getRange(filaRegistro, 8).setValue(horaAhora);
        sheet.getRange(filaRegistro, 9).setValue(ubicacion);
        Logger.log('✅ Comida-In registrada');
        return { exito: true, mensaje: 'Comida-In registrada a las ' + horaAhora };
      
      case 'comida-out':
        if (filaRegistro === -1) {
          return { exito: false, mensaje: "Primero registra la Entrada." };
        }
        if (!sheet.getRange(filaRegistro, 8).getValue()) {
          return { exito: false, mensaje: "Primero registra Comida-In." };
        }
        if (sheet.getRange(filaRegistro, 10).getValue()) {
          return { exito: false, mensaje: "Ya registraste Comida-Out." };
        }
        sheet.getRange(filaRegistro, 10).setValue(horaAhora);
        sheet.getRange(filaRegistro, 11).setValue(ubicacion);
        Logger.log('✅ Comida-Out registrada');
        return { exito: true, mensaje: 'Comida-Out registrada a las ' + horaAhora };
      
      case 'salida':
        if (filaRegistro === -1) {
          return { exito: false, mensaje: "Primero registra la Entrada." };
        }
        if (sheet.getRange(filaRegistro, 8).getValue() && !sheet.getRange(filaRegistro, 10).getValue()) {
          return { exito: false, mensaje: "Registra Comida-Out antes de salir." };
        }
        let fotoSalidaUrl = '';
        if (data.foto && data.foto.length > 100) {
          fotoSalidaUrl = guardarFoto(data.foto, data.usuario + '_salida_' + Date.now());
        }
        sheet.getRange(filaRegistro, 12).setValue(fechaHoyDisplay);
        sheet.getRange(filaRegistro, 13).setValue(horaAhora);
        sheet.getRange(filaRegistro, 14).setValue(ubicacion);
        sheet.getRange(filaRegistro, 15).setValue(fotoSalidaUrl);
        const tiempos = calcularTiempos(filaRegistro, sheet);
        sheet.getRange(filaRegistro, 16).setValue(tiempos.productivo);
        sheet.getRange(filaRegistro, 17).setValue(tiempos.total);
        Logger.log('✅ Salida registrada');
        return { exito: true, mensaje: 'Salida registrada. Total: ' + tiempos.total + ' | Productivo: ' + tiempos.productivo };
      
      default:
        return { exito: false, mensaje: "Tipo no válido" };
    }
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    return { exito: false, mensaje: "Error: " + error.message };
  }
}

// ==================== EMPRESA ====================
function guardarEmpresa(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Empresa');
  if (!sheet) return { exito: false, mensaje: "No existe hoja 'Empresa'. Créala con encabezados: Nombre, Slogan, Direccion, Latitud, Longitud" };
  
  sheet.getRange(2, 1).setValue(data.nombre || '');
  sheet.getRange(2, 2).setValue(data.slogan || '');
  sheet.getRange(2, 3).setValue(data.direccion || '');
  sheet.getRange(2, 4).setValue(data.latitud || '');
  sheet.getRange(2, 5).setValue(data.longitud || '');
  
  return { exito: true, mensaje: 'Datos de la empresa guardados' };
}

function obtenerEmpresa() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Empresa');
  if (!sheet) return { exito: false, mensaje: "No existe hoja 'Empresa'" };
  
  const datos = sheet.getRange(2, 1, 1, 5).getValues()[0];
  return {
    exito: true,
    empresa: {
      nombre: String(datos[0] || '').trim(),
      slogan: String(datos[1] || '').trim(),
      direccion: String(datos[2] || '').trim(),
      latitud: String(datos[3] || '').trim(),
      longitud: String(datos[4] || '').trim()
    }
  };
}

// ==================== EMPLEADOS ====================
function crearEmpleado(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(data.usuario).trim()) {
      return { exito: false, mensaje: "El usuario ya existe" };
    }
  }
  sheet.appendRow([data.usuario, data.password, data.nombre, data.area, data.cargo, data.perfil, data.celular || '', data.correo || '']);
  return { exito: true, mensaje: 'Empleado "' + data.nombre + '" creado' };
}

function listarEmpleados() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const datos = sheet.getDataRange().getValues();
  const empleados = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] && String(datos[i][0]).trim() !== '') {
      empleados.push({
        usuario: String(datos[i][0]).trim(),
        password: String(datos[i][1]).trim(),
        nombre: String(datos[i][2]).trim(),
        area: String(datos[i][3]).trim(),
        cargo: String(datos[i][4]).trim(),
        perfil: String(datos[i][5]).trim(),
        celular: String(datos[i][6] || '').trim(),
        correo: String(datos[i][7] || '').trim()
      });
    }
  }
  return { exito: true, empleados: empleados, total: empleados.length };
}

function eliminarEmpleado(usuario) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(usuario).trim()) {
      sheet.deleteRow(i + 1);
      return { exito: true, mensaje: 'Empleado "' + usuario + '" eliminado' };
    }
  }
  return { exito: false, mensaje: "Empleado no encontrado" };
}

function editarEmpleado(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(data.usuario).trim()) {
      const fila = i + 1;
      if (data.area !== undefined) sheet.getRange(fila, 4).setValue(data.area);
      if (data.cargo !== undefined) sheet.getRange(fila, 5).setValue(data.cargo);
      if (data.celular !== undefined) sheet.getRange(fila, 7).setValue(data.celular);
      if (data.correo !== undefined) sheet.getRange(fila, 8).setValue(data.correo);
      return { exito: true, mensaje: 'Datos de "' + data.usuario + '" actualizados' };
    }
  }
  return { exito: false, mensaje: 'Empleado no encontrado' };
}

function cambiarPassword(usuario, passwordActual, passwordNueva) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(usuario).trim()) {
      if (String(datos[i][1]).trim() !== String(passwordActual).trim()) {
        return { exito: false, mensaje: "La contraseña actual es incorrecta" };
      }
      sheet.getRange(i + 1, 2).setValue(passwordNueva);
      return { exito: true, mensaje: "Contraseña actualizada" };
    }
  }
  return { exito: false, mensaje: "Usuario no encontrado" };
}

// ==================== ASISTENCIAS ====================
function listarAsistencias() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Asistencias');
  const datos = sheet.getDataRange().getValues();
  const registros = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] && String(datos[i][1]).trim() !== '') {
      registros.push({
        id: String(datos[i][0]).trim(),
        usuario: String(datos[i][1]).trim(),
        nombre: String(datos[i][2]).trim(),
        fechaEntrada: formatearFecha(datos[i][3]),
        horaEntrada: formatearHora(datos[i][4]),
        ubicacionEntrada: String(datos[i][5]).trim(),
        fotoEntrada: String(datos[i][6]).trim(),
        horaIniRef: formatearHora(datos[i][7]),
        ubicacionIniRef: String(datos[i][8]).trim(),
        horaFinRef: formatearHora(datos[i][9]),
        ubicacionFinRef: String(datos[i][10]).trim(),
        fechaSalida: formatearFecha(datos[i][11]),
        horaSalida: formatearHora(datos[i][12]),
        ubicacionSalida: String(datos[i][13]).trim(),
        fotoSalida: String(datos[i][14]).trim(),
        tiempoProductivo: String(datos[i][15]).trim(),
        tiempoTotal: String(datos[i][16]).trim()
      });
    }
  }
  return { exito: true, registros: registros, total: registros.length };
}

// ==================== AUXILIARES ====================
function formatearFecha(valor) {
  if (!valor) return '';
  if (valor instanceof Date) return Utilities.formatDate(valor, 'America/Caracas', 'dd/MM/yyyy');
  return String(valor).trim();
}

function formatearHora(valor) {
  if (!valor) return '';
  if (valor instanceof Date) return Utilities.formatDate(valor, 'America/Caracas', 'HH:mm:ss');
  return String(valor).trim();
}

function calcularTiempos(fila, sheet) {
  const entrada = sheet.getRange(fila, 5).getValue();
  const salida = sheet.getRange(fila, 13).getValue();
  if (!entrada || !salida) return { productivo: 'N/A', total: 'N/A' };
  const e = parsearHora(entrada);
  const s = parsearHora(salida);
  let totalMs = s - e;
  let productivoMs = totalMs;
  const iniRef = sheet.getRange(fila, 8).getValue();
  const finRef = sheet.getRange(fila, 10).getValue();
  if (iniRef && finRef) productivoMs = totalMs - (parsearHora(finRef) - parsearHora(iniRef));
  return { total: formatearDuracion(totalMs), productivo: formatearDuracion(productivoMs) };
}

function parsearHora(valor) {
  if (valor instanceof Date) return valor;
  const partes = String(valor).trim().split(':');
  const hoy = new Date();
  hoy.setHours(parseInt(partes[0]), parseInt(partes[1]), parseInt(partes[2]) || 0, 0);
  return hoy;
}

function formatearDuracion(ms) {
  if (ms < 0) return '00:00:00';
  const h = Math.floor(ms / 3600000);
  const m = Math.floor((ms % 3600000) / 60000);
  const s = Math.floor((ms % 60000) / 1000);
  return pad2(h) + ':' + pad2(m) + ':' + pad2(s);
}

function pad2(n) { return n < 10 ? '0' + n : '' + n; }

function guardarFoto(base64, nombre) {
  try {
    const carpeta = DriveApp.getFoldersByName('Fotos_Asistencia').hasNext() 
      ? DriveApp.getFoldersByName('Fotos_Asistencia').next() 
      : DriveApp.createFolder('Fotos_Asistencia');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64.split(',')[1]), 'image/jpeg', nombre + '.jpg');
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch (error) {
    return 'Error: ' + error.message;
  }
}