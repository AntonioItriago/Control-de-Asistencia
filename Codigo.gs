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
    case 'listar_asistencias':
      respuesta = listarAsistencias();
      break;
    case 'cambiar_password':
      respuesta = cambiarPassword(data.usuario, data.passwordActual, data.passwordNueva);
      break;
    case 'forzar_registro':
      respuesta = forzarRegistro(data);
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
        rol: String(datos[i][5]).trim(),
        nombre: String(datos[i][2]).trim(),
        area: String(datos[i][3]).trim(),
        cargo: String(datos[i][4]).trim()
      };
    }
  }
  return { exito: false, mensaje: "Usuario o contraseña incorrectos" };
}

// ==================== REGISTRAR ASISTENCIA (con validaciones de secuencia) ====================
function registrarAsistencia(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Asistencias');
  if (!sheet) return { exito: false, mensaje: "No existe hoja 'Asistencias'" };
  
  const datos = sheet.getDataRange().getValues();
  const ahora = new Date();
  const fechaHoy = Utilities.formatDate(ahora, 'America/Caracas', 'dd/MM/yyyy');
  const horaAhora = Utilities.formatDate(ahora, 'America/Caracas', 'HH:mm:ss');
  const ubicacion = `${data.lat.toFixed(6)}, ${data.lng.toFixed(6)}`;
  
  // Buscar registro de HOY para este usuario
  let filaRegistro = -1;
  for (let i = 1; i < datos.length; i++) {
    const usuarioFila = String(datos[i][1]).trim();
    const fechaFila = String(datos[i][3]).trim();
    if (usuarioFila === data.usuario && fechaFila === fechaHoy) {
      filaRegistro = i + 1;
      break;
    }
  }
  
  // Buscar registro de AYER para validar que cerró jornada
  let filaAyer = -1;
  const ayer = new Date(ahora);
  ayer.setDate(ayer.getDate() - 1);
  const fechaAyer = Utilities.formatDate(ayer, 'America/Caracas', 'dd/MM/yyyy');
  for (let i = 1; i < datos.length; i++) {
    const usuarioFila = String(datos[i][1]).trim();
    const fechaFila = String(datos[i][3]).trim();
    if (usuarioFila === data.usuario && fechaFila === fechaAyer) {
      filaAyer = i + 1;
      break;
    }
  }
  
  try {
    switch (data.tipo) {
      case 'entrada':
        // VALIDACIÓN: No puede tener entrada hoy
        if (filaRegistro !== -1) {
          return { exito: false, mensaje: "Ya registraste entrada hoy. Usa los otros botones." };
        }
        // VALIDACIÓN: Debe haber cerrado jornada de ayer
        if (filaAyer !== -1) {
          const horaSalidaAyer = sheet.getRange(filaAyer, 13).getValue(); // Columna M
          if (!horaSalidaAyer) {
            return { exito: false, mensaje: "No puedes iniciar hoy porque no cerraste la jornada de ayer. Contacta al administrador." };
          }
        }
        
        let fotoEntradaUrl = '';
        if (data.foto && data.foto.length > 100) {
          fotoEntradaUrl = guardarFoto(data.foto, `${data.usuario}_entrada_${Date.now()}`);
        }
        
        const nuevoId = 'AS-' + ahora.getTime();
        sheet.appendRow([
          nuevoId, data.usuario, data.nombre || data.usuario,
          fechaHoy, horaAhora, ubicacion, fotoEntradaUrl,
          '', '', '', '', '', '', '', '', '', '', ''
        ]);
        return { exito: true, mensaje: `Entrada registrada a las ${horaAhora}` };
      
      case 'comida-in':
        if (filaRegistro === -1) return { exito: false, mensaje: "Primero debes registrar la Entrada." };
        // Validar que no haya ya comida-in
        const horaIniRef = sheet.getRange(filaRegistro, 8).getValue();
        if (horaIniRef) return { exito: false, mensaje: "Ya registraste Comida-In hoy. Debes registrar Comida-Out primero." };
        sheet.getRange(filaRegistro, 8).setValue(horaAhora);
        sheet.getRange(filaRegistro, 9).setValue(ubicacion);
        return { exito: true, mensaje: `Inicio de descanso registrado a las ${horaAhora}` };
      
      case 'comida-out':
        if (filaRegistro === -1) return { exito: false, mensaje: "Primero debes registrar la Entrada." };
        const horaIniRef2 = sheet.getRange(filaRegistro, 8).getValue();
        if (!horaIniRef2) return { exito: false, mensaje: "Primero debes registrar Comida-In." };
        const horaFinRef = sheet.getRange(filaRegistro, 10).getValue();
        if (horaFinRef) return { exito: false, mensaje: "Ya registraste Comida-Out hoy." };
        sheet.getRange(filaRegistro, 10).setValue(horaAhora);
        sheet.getRange(filaRegistro, 11).setValue(ubicacion);
        return { exito: true, mensaje: `Fin de descanso registrado a las ${horaAhora}` };
      
      case 'salida':
        if (filaRegistro === -1) return { exito: false, mensaje: "Primero debes registrar la Entrada." };
        // Validar que haya comida-out (si registró comida-in)
        const hiRef = sheet.getRange(filaRegistro, 8).getValue();
        const hfRef = sheet.getRange(filaRegistro, 10).getValue();
        if (hiRef && !hfRef) {
          return { exito: false, mensaje: "Registraste Comida-In pero no Comida-Out. Completa el registro primero." };
        }
        
        let fotoSalidaUrl = '';
        if (data.foto && data.foto.length > 100) {
          fotoSalidaUrl = guardarFoto(data.foto, `${data.usuario}_salida_${Date.now()}`);
        }
        
        sheet.getRange(filaRegistro, 12).setValue(fechaHoy);
        sheet.getRange(filaRegistro, 13).setValue(horaAhora);
        sheet.getRange(filaRegistro, 14).setValue(ubicacion);
        sheet.getRange(filaRegistro, 15).setValue(fotoSalidaUrl);
        
        const tiempos = calcularTiempos(filaRegistro, sheet);
        sheet.getRange(filaRegistro, 16).setValue(tiempos.productivo);
        sheet.getRange(filaRegistro, 17).setValue(tiempos.total);
        
        return { exito: true, mensaje: `Salida registrada. Total: ${tiempos.total} | Productivo: ${tiempos.productivo}` };
      
      default:
        return { exito: false, mensaje: "Tipo no válido" };
    }
  } catch (error) {
    return { exito: false, mensaje: "Error: " + error.message };
  }
}

// ==================== FORZAR REGISTRO (ADMIN) ====================
function forzarRegistro(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Asistencias');
  const datos = sheet.getDataRange().getValues();
  
  let filaRegistro = -1;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim() === String(data.usuario).trim() && 
        String(datos[i][3]).trim() === String(data.fecha).trim()) {
      filaRegistro = i + 1;
      break;
    }
  }
  
  if (filaRegistro === -1) {
    return { exito: false, mensaje: "No se encontró registro para ese usuario en esa fecha" };
  }
  
  // data.columna = número de columna (8=HoraIniRef, 10=HoraFinRef, 13=HoraSalida, etc.)
  sheet.getRange(filaRegistro, data.columna).setValue(data.valor);
  return { exito: true, mensaje: `Campo actualizado a ${data.valor}` };
}

// ==================== CALCULAR TIEMPOS ====================
function calcularTiempos(fila, sheet) {
  const horaEntrada = sheet.getRange(fila, 5).getValue();
  const horaIniRef = sheet.getRange(fila, 8).getValue();
  const horaFinRef = sheet.getRange(fila, 10).getValue();
  const horaSalida = sheet.getRange(fila, 13).getValue();
  
  if (!horaEntrada || !horaSalida) return { productivo: 'N/A', total: 'N/A' };
  
  const entrada = parsearHora(horaEntrada);
  const salida = parsearHora(horaSalida);
  const totalMs = salida - entrada;
  
  let productivoMs = totalMs;
  if (horaIniRef && horaFinRef) {
    productivoMs = totalMs - (parsearHora(horaFinRef) - parsearHora(horaIniRef));
  }
  
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
  return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
}

// ==================== GUARDAR FOTO EN DRIVE ====================
function guardarFoto(base64, nombre) {
  try {
    const carpeta = DriveApp.getFoldersByName('Fotos_Asistencia').hasNext() 
      ? DriveApp.getFoldersByName('Fotos_Asistencia').next() 
      : DriveApp.createFolder('Fotos_Asistencia');
    const base64Data = base64.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', nombre + '.jpg');
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch (error) {
    return 'Error: ' + error.message;
  }
}

// ==================== CRUD EMPLEADOS ====================
function crearEmpleado(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(data.usuario).trim()) {
      return { exito: false, mensaje: "El usuario ya existe" };
    }
  }
  sheet.appendRow([data.usuario, data.password, data.nombre, data.area, data.cargo, data.perfil]);
  return { exito: true, mensaje: `Empleado "${data.nombre}" creado` };
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
        perfil: String(datos[i][5]).trim()
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
      return { exito: true, mensaje: `Empleado "${usuario}" eliminado` };
    }
  }
  return { exito: false, mensaje: "Empleado no encontrado" };
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

function listarAsistencias() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Asistencias');
  const datos = sheet.getDataRange().getValues();
  const registros = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1]) {
      registros.push({
        id: String(datos[i][0]).trim(),
        usuario: String(datos[i][1]).trim(),
        nombre: String(datos[i][2]).trim(),
        fechaEntrada: String(datos[i][3]).trim(),
        horaEntrada: String(datos[i][4]).trim(),
        ubicacionEntrada: String(datos[i][5]).trim(),
        fotoEntrada: String(datos[i][6]).trim(),
        horaIniRef: String(datos[i][7]).trim(),
        ubicacionIniRef: String(datos[i][8]).trim(),
        horaFinRef: String(datos[i][9]).trim(),
        ubicacionFinRef: String(datos[i][10]).trim(),
        fechaSalida: String(datos[i][11]).trim(),
        horaSalida: String(datos[i][12]).trim(),
        ubicacionSalida: String(datos[i][13]).trim(),
        fotoSalida: String(datos[i][14]).trim(),
        tiempoProductivo: String(datos[i][15]).trim(),
        tiempoTotal: String(datos[i][16]).trim()
      });
    }
  }
  return { exito: true, registros: registros, total: registros.length };
}