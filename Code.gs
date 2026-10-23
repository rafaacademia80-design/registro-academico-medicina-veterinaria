/**
 * Registro y actualización de datos académicos
 * Carrera: Medicina Veterinaria
 * La lista maestra SOLO autoriza: cédula + periodo + subproyecto + sección.
 * Los datos personales son declarados directamente por el estudiante.
 */
const CFG = {
  ROSTER: 'Lista_Clases',
  RECORDS: 'Registros',
  AUDIT: 'Auditoria',
  PERIODOS: ['2026-II', '2027-I', '2027-II'],
  SECCIONES: ['D02', 'D03', 'D04', 'D05', 'D06', 'D07', 'D08', 'D09'],
  SUBPROYECTOS: [
    'Sistemas de Producción Animal I',
    'Sistemas de Producción Animal II',
    'Manejo de Sistemas de Producción Animal I',
    'Manejo de Sistemas de Producción Animal II'
  ],
  SEMESTRES: {
    'Sistemas de Producción Animal I': 'Primer semestre',
    'Sistemas de Producción Animal II': 'Segundo semestre',
    'Manejo de Sistemas de Producción Animal I': 'Tercer semestre',
    'Manejo de Sistemas de Producción Animal II': 'Cuarto semestre'
  },
  OTP_TTL_SECONDS: 300,
  OTP_COOLDOWN_SECONDS: 60,
  OTP_MAX_ATTEMPTS: 5
};

const HEADERS = {
  ROSTER: ['CEDULA', 'PERIODO', 'SUBPROYECTO', 'SECCION'],
  RECORDS: [
    'FECHA_REGISTRO', 'ULTIMA_ACTUALIZACION', 'CEDULA', 'NOMBRE_COMPLETO',
    'EDAD', 'CORREO_CONTACTO', 'TELEFONO', 'LUGAR_DE_RESIDENCIA', 'SUBPROYECTO',
    'SEMESTRE', 'SECCION', 'PERIODO', 'OBSERVACIONES',
    'CORREO_VERIFICADO', 'ESTADO_CORREO'
  ],
  AUDIT: ['FECHA_HORA', 'CEDULA', 'PERIODO', 'SUBPROYECTO', 'SECCION', 'OPERACION', 'RESULTADO']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Registro académico | Medicina Veterinaria')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Ejecutar una vez desde el editor vinculado a la hoja de cálculo. */
function prepararSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Vincule este proyecto de Apps Script a una hoja de cálculo primero.');
  asegurarHoja_(ss, CFG.ROSTER, HEADERS.ROSTER);
  asegurarHoja_(ss, CFG.RECORDS, HEADERS.RECORDS);
  asegurarHoja_(ss, CFG.AUDIT, HEADERS.AUDIT);
  const roster = ss.getSheetByName(CFG.ROSTER);
  roster.setFrozenRows(1);
  roster.getRange('A:A').setNumberFormat('@');
  const records = ss.getSheetByName(CFG.RECORDS);
  records.setFrozenRows(1);
  const hm = mapaEncabezados_(records.getRange(1, 1, 1, records.getLastColumn()).getDisplayValues()[0]);
  if (hm.CEDULA !== undefined) records.getRange(2, hm.CEDULA + 1, Math.max(records.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  SpreadsheetApp.flush();
  return 'Sistema preparado. La Lista_Clases debe contener únicamente CEDULA, PERIODO, SUBPROYECTO y SECCION. Revise la guía si ya tenía instalada una versión anterior.';
}

/** Solicita autorización al servicio de correo; no envía ningún mensaje. */
function autorizarServicios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Vincule este proyecto de Apps Script a una hoja de cálculo primero.');
  const cuota = MailApp.getRemainingDailyQuota();
  return 'Servicio de correo autorizado. Destinatarios disponibles hoy según Apps Script: ' + cuota;
}

/** Comprueba únicamente que la asignación académica exista en la lista maestra. */
function validarAcceso(datos) {
  try {
    const acceso = normalizarAcceso_(datos || {});
    const autorizado = buscarAutorizacion_(acceso);
    if (!autorizado.ok) return { ok: false, message: autorizado.message };
    const coincidencias = buscarRegistros_(acceso);
    if (coincidencias.length > 1) {
      return { ok: false, message: 'Existen registros duplicados para esta cédula, periodo y subproyecto. Contacte al docente antes de continuar.' };
    }
    return {
      ok: true,
      existeRegistro: coincidencias.length === 1,
      message: coincidencias.length === 1
        ? 'Asignación autorizada. Se detectó un registro previo; para actualizarlo deberá verificar el correo asociado a ese registro.'
        : 'Asignación autorizada. Complete sus datos reales y actuales; enviaremos un código al correo que indique para confirmar que puede recibir mensajes allí.'
    };
  } catch (e) {
    return { ok: false, message: e && e.message ? e.message : 'No se pudo validar la asignación académica.' };
  }
}

/**
 * En un alta inicial, el código se envía al correo que el estudiante declara.
 * En una actualización, se envía al correo de autenticación ya asociado al registro,
 * no al correo nuevo que se escribe en el formulario.
 */
function solicitarCodigo(datos) {
  try {
    const acceso = normalizarAcceso_(datos || {});
    const autorizado = buscarAutorizacion_(acceso);
    if (!autorizado.ok) return { ok: false, message: autorizado.message };

    const coincidencias = buscarRegistros_(acceso);
    if (coincidencias.length > 1) {
      return { ok: false, message: 'Existen registros duplicados para esta cédula, periodo y subproyecto. El docente debe depurarlos antes de continuar.' };
    }

    const esActualizacion = coincidencias.length === 1;
    let correoDestino;
    if (esActualizacion) {
      const fila = coincidencias[0].registro;
      correoDestino = String(fila.CORREO_VERIFICADO || '').trim().toLowerCase();
      if (!correoValido_(correoDestino)) {
        return { ok: false, message: 'Este registro no tiene un correo de autenticación verificado. Contacte al docente para restablecer el acceso y actualizar el correo.' };
      }
    } else {
      correoDestino = String(datos.correo || '').trim().toLowerCase();
      if (!correoValido_(correoDestino)) return { ok: false, message: 'Escriba un correo electrónico válido antes de solicitar el código.' };
      if (String(datos.correo || '').trim().toLowerCase() !== String(datos.correoConfirmacion || '').trim().toLowerCase()) {
        return { ok: false, message: 'El correo y su confirmación no coinciden.' };
      }
    }

    const cache = CacheService.getScriptCache();
    const otpKey = claveOtp_(acceso);
    const rateKey = 'rate_' + otpKey;
    const mailRateKey = 'rate_mail_' + hashTexto_(correoDestino).slice(0, 32);
    if (cache.get(rateKey) || cache.get(mailRateKey)) {
      return { ok: false, message: 'Ya se solicitó un código recientemente para esta solicitud o correo. Espere un minuto antes de pedir otro.' };
    }

    const codigo = String(Math.floor(100000 + Math.random() * 900000));
    const asunto = 'Código de verificación de registro académico';
    const cuerpo = [
      'Su código para el registro o actualización académica es: ' + codigo,
      '',
      'El código vence en 5 minutos y solo puede utilizarse una vez.',
      'Periodo: ' + acceso.PERIODO,
      'Subproyecto: ' + acceso.SUBPROYECTO,
      'Sección: ' + acceso.SECCION,
      '',
      'Si usted no inició esta solicitud, ignore este mensaje y comuníquese con el docente responsable.',
      'No comparta este código con otras personas.'
    ].join('\n');

    try {
      MailApp.sendEmail(correoDestino, asunto, cuerpo, { name: 'Registro Académico de Medicina Veterinaria' });
    } catch (mailError) {
      return { ok: false, message: 'No se pudo enviar el código. Compruebe la autorización y la cuota diaria de correo de Google Apps Script.' };
    }

    const paquete = {
      codigo: codigo,
      intentos: 0,
      creadoEn: Date.now(),
      correoDestino: correoDestino,
      tipo: esActualizacion ? 'ACTUALIZACION' : 'REGISTRO'
    };
    cache.put(otpKey, JSON.stringify(paquete), CFG.OTP_TTL_SECONDS);
    cache.put(rateKey, '1', CFG.OTP_COOLDOWN_SECONDS);
    cache.put(mailRateKey, '1', CFG.OTP_COOLDOWN_SECONDS);
    return {
      ok: true,
      tipo: paquete.tipo,
      correo: enmascararCorreo_(correoDestino),
      message: esActualizacion
        ? 'Código enviado al correo de autenticación asociado al registro previo.'
        : 'Código enviado al correo que indicó. La verificación confirma que puede recibir mensajes en esa dirección.'
    };
  } catch (e) {
    console.error(e);
    return { ok: false, message: e && e.message ? e.message : 'No fue posible solicitar el código de verificación.' };
  }
}

function guardarRegistro(datos) {
  const lock = LockService.getScriptLock();
  try {
    const d = validarFormulario_(datos || {});
    const acceso = { CEDULA: d.CEDULA, PERIODO: d.PERIODO, SUBPROYECTO: d.SUBPROYECTO, SECCION: d.SECCION };
    lock.waitLock(10000);
    try {
      const autorizado = buscarAutorizacion_(acceso);
      if (!autorizado.ok) return { ok: false, message: autorizado.message };

      const coincidencias = buscarRegistros_(acceso);
      if (coincidencias.length > 1) {
        return { ok: false, message: 'Se detectaron duplicados. El docente debe depurarlos antes de actualizar los datos.' };
      }
      const existe = coincidencias.length === 1;
      const cache = CacheService.getScriptCache();
      const otpKey = claveOtp_(acceso);
      const otpRaw = cache.get(otpKey);
      if (!otpRaw) return { ok: false, message: 'El código no existe o venció. Solicite uno nuevo.' };
      const otpData = JSON.parse(otpRaw);
      if (!otpData.creadoEn || Date.now() - Number(otpData.creadoEn) > CFG.OTP_TTL_SECONDS * 1000) {
        cache.remove(otpKey);
        return { ok: false, message: 'El código venció. Solicite uno nuevo.' };
      }
      const codigoIngresado = String(d.CODIGO_VERIFICACION || '').replace(/\D/g, '');
      if (codigoIngresado !== otpData.codigo) {
        otpData.intentos = Number(otpData.intentos || 0) + 1;
        if (otpData.intentos >= CFG.OTP_MAX_ATTEMPTS) {
          cache.remove(otpKey);
          return { ok: false, message: 'Se superó el número de intentos. Solicite un nuevo código.' };
        }
        const segundos = Math.max(1, Math.ceil((Number(otpData.creadoEn) + CFG.OTP_TTL_SECONDS * 1000 - Date.now()) / 1000));
        cache.put(otpKey, JSON.stringify(otpData), segundos);
        return { ok: false, message: 'Código incorrecto. Revíselo e inténtelo nuevamente.' };
      }

      if (existe) {
        const correoAccesoActual = String(coincidencias[0].registro.CORREO_VERIFICADO || '').trim().toLowerCase();
        if (otpData.tipo !== 'ACTUALIZACION' || correoAccesoActual !== String(otpData.correoDestino || '').toLowerCase()) {
          cache.remove(otpKey);
          return { ok: false, message: 'El registro o el correo de autenticación cambió durante el proceso. Solicite un código nuevo.' };
        }
      } else {
        if (otpData.tipo !== 'REGISTRO' || String(d.CORREO_CONTACTO).toLowerCase() !== String(otpData.correoDestino || '').toLowerCase()) {
          cache.remove(otpKey);
          return { ok: false, message: 'El correo cambió después de solicitar el código. Solicite un nuevo código para el correo confirmado.' };
        }
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(CFG.RECORDS);
      if (!sh) return { ok: false, message: 'No existe la hoja Registros. Ejecute prepararSistema() desde Apps Script.' };
      const all = sh.getDataRange().getValues();
      const headers = all[0] || [];
      const hm = mapaEncabezados_(headers);
      if (!HEADERS.RECORDS.every(h => Object.prototype.hasOwnProperty.call(hm, h))) {
        return { ok: false, message: 'Los encabezados de Registros no coinciden con la versión actual. Use una hoja nueva o revise la guía de migración.' };
      }

      const ahora = new Date();
      const filaExistente = existe ? coincidencias[0].fila : null;
      const fechaRegistro = filaExistente ? (sh.getRange(filaExistente, hm.FECHA_REGISTRO + 1).getValue() || ahora) : ahora;
      const correoAutenticacion = existe
        ? String(coincidencias[0].registro.CORREO_VERIFICADO || '').trim().toLowerCase()
        : d.CORREO_CONTACTO;
      const estadoCorreo = d.CORREO_CONTACTO === correoAutenticacion ? 'VERIFICADO' : 'CAMBIO_PENDIENTE_DE_VERIFICACION';
      const registro = {
        FECHA_REGISTRO: fechaRegistro,
        ULTIMA_ACTUALIZACION: ahora,
        CEDULA: d.CEDULA,
        NOMBRE_COMPLETO: d.NOMBRE_COMPLETO,
        EDAD: d.EDAD,
        CORREO_CONTACTO: d.CORREO_CONTACTO,
        TELEFONO: d.TELEFONO,
        RESIDENCIA: d.RESIDENCIA,
        SUBPROYECTO: d.SUBPROYECTO,
        SEMESTRE: CFG.SEMESTRES[d.SUBPROYECTO],
        SECCION: d.SECCION,
        PERIODO: d.PERIODO,
        OBSERVACIONES: d.OBSERVACIONES,
        CORREO_VERIFICADO: correoAutenticacion,
        ESTADO_CORREO: estadoCorreo
      };
      const nuevaFila = headers.map(h => registro[String(h).trim().toUpperCase()] !== undefined ? registro[String(h).trim().toUpperCase()] : '');
      if (filaExistente) sh.getRange(filaExistente, 1, 1, nuevaFila.length).setValues([nuevaFila]);
      else sh.appendRow(nuevaFila);

      cache.remove(otpKey);
      registrarAuditoria_(ss, d, filaExistente ? 'ACTUALIZACION' : 'REGISTRO', 'OK');
      SpreadsheetApp.flush();
      return {
        ok: true,
        operacion: filaExistente ? 'actualizado' : 'registrado',
        message: filaExistente
          ? (estadoCorreo === 'VERIFICADO'
            ? 'Los datos se actualizaron correctamente.'
            : 'Los datos se actualizaron. El nuevo correo de contacto quedó pendiente de verificación; el correo de autenticación anterior se conserva para proteger futuras actualizaciones.')
          : 'El registro se guardó correctamente y el correo indicado fue verificado mediante código.'
      };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    console.error(e);
    return { ok: false, message: e && e.message ? e.message : 'No se pudo guardar el registro. Revise los datos e inténtelo nuevamente.' };
  }
}

function normalizarAcceso_(d) {
  const tipo = String(d.tipoCedula || '').trim().toUpperCase();
  const numero = String(d.numeroCedula || '').replace(/\D/g, '');
  if (!['V', 'E'].includes(tipo) || !/^\d{5,9}$/.test(numero)) {
    throw new Error('Ingrese un tipo de cédula válido (V o E) y entre 5 y 9 dígitos, sin puntos ni guiones.');
  }
  const acceso = {
    CEDULA: tipo + '-' + numero,
    PERIODO: String(d.periodo || '').trim().toUpperCase(),
    SUBPROYECTO: String(d.subproyecto || '').trim(),
    SECCION: String(d.seccion || '').trim().toUpperCase()
  };
  if (!CFG.PERIODOS.includes(acceso.PERIODO)) throw new Error('Seleccione un periodo lectivo válido.');
  if (!CFG.SUBPROYECTOS.includes(acceso.SUBPROYECTO)) throw new Error('Seleccione uno de los cuatro subproyectos autorizados.');
  if (!CFG.SECCIONES.includes(acceso.SECCION)) throw new Error('Seleccione una sección válida.');
  return acceso;
}

function buscarAutorizacion_(acceso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss && ss.getSheetByName(CFG.ROSTER);
  if (!sh) return { ok: false, message: 'El responsable académico aún no ha configurado la lista de clases.' };
  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return { ok: false, message: 'La lista de clases está vacía. Contacte al responsable académico.' };
  const hm = mapaEncabezados_(values[0]);
  if (!HEADERS.ROSTER.every(h => Object.prototype.hasOwnProperty.call(hm, h))) {
    return { ok: false, message: 'Lista_Clases debe tener los encabezados CEDULA, PERIODO, SUBPROYECTO y SECCION.' };
  }
  const matches = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const ci = normalizarCedulaTexto_(row[hm.CEDULA]);
    const periodo = String(row[hm.PERIODO] || '').trim().toUpperCase();
    const sub = String(row[hm.SUBPROYECTO] || '').trim();
    const seccion = String(row[hm.SECCION] || '').trim().toUpperCase();
    if (ci === acceso.CEDULA && periodo === acceso.PERIODO && normalizarTexto_(sub) === normalizarTexto_(acceso.SUBPROYECTO) && seccion === acceso.SECCION) {
      matches.push(i + 1);
    }
  }
  if (matches.length === 0) return { ok: false, message: 'La combinación de cédula, periodo, subproyecto y sección no aparece en la lista autorizada. Revísela o contacte al docente.' };
  if (matches.length > 1) return { ok: false, message: 'La lista de clases tiene duplicada esta combinación. El docente debe depurarla antes de continuar.' };
  return { ok: true, fila: matches[0] };
}

function buscarRegistros_(acceso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss && ss.getSheetByName(CFG.RECORDS);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const hm = mapaEncabezados_(vals[0] || []);
  const requeridos = ['CEDULA', 'PERIODO', 'SUBPROYECTO', 'CORREO_CONTACTO', 'CORREO_VERIFICADO', 'FECHA_REGISTRO'];
  if (!requeridos.every(h => Object.prototype.hasOwnProperty.call(hm, h))) {
    throw new Error('La hoja Registros tiene una estructura de una versión anterior. Consulte la guía antes de utilizar datos reales.');
  }
  const encontrados = [];
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    if (normalizarCedulaTexto_(row[hm.CEDULA]) === acceso.CEDULA &&
        String(row[hm.PERIODO] || '').trim().toUpperCase() === acceso.PERIODO &&
        normalizarTexto_(row[hm.SUBPROYECTO]) === normalizarTexto_(acceso.SUBPROYECTO)) {
      const registro = {};
      Object.keys(hm).forEach(k => { registro[k] = row[hm[k]] === undefined ? '' : row[hm[k]]; });
      encontrados.push({ fila: i + 1, registro: registro });
    }
  }
  return encontrados;
}

function validarFormulario_(d) {
  const acceso = normalizarAcceso_(d);
  const edad = Number(d.edad);
  if (!Number.isInteger(edad) || edad < 15 || edad > 90) throw new Error('La edad debe ser un número entero entre 15 y 90.');
  const email = String(d.correo || '').trim().toLowerCase();
  const emailConfirmacion = String(d.correoConfirmacion || '').trim().toLowerCase();
  if (!correoValido_(email)) throw new Error('Ingrese un correo electrónico válido.');
  if (email !== emailConfirmacion) throw new Error('El correo electrónico y su confirmación no coinciden.');
  const residencia = limpiarTexto_(d.residencia, 120);
  if (residencia.length < 3) throw new Error('Indique municipio y estado de residencia.');
  const telefono = String(d.telefono || '').trim().replace(/\s+/g, ' ');
  const digitosTelefono = telefono.replace(/\D/g, '');
  if (digitosTelefono.length < 7 || digitosTelefono.length > 15 || !/^[0-9+()\-. ]+$/.test(telefono)) {
    throw new Error('Ingrese un teléfono válido, con entre 7 y 15 dígitos. Puede incluir código de país, espacios, paréntesis o guiones.');
  }
  const codigo = String(d.codigoVerificacion || '').replace(/\D/g, '');
  if (!/^\d{6}$/.test(codigo)) throw new Error('El código de verificación debe tener 6 dígitos.');
  if (d.consentimiento !== true && String(d.consentimiento) !== 'true') throw new Error('Debe confirmar la declaración de veracidad y uso académico de los datos.');

  const primerApellido = validarCampoNombre_(d.primerApellido, 'primer apellido', true);
  const segundoApellido = validarCampoNombre_(d.segundoApellido, 'segundo apellido', false);
  const primerNombre = validarCampoNombre_(d.primerNombre, 'primer nombre', true);
  const segundoNombre = validarCampoNombre_(d.segundoNombre, 'segundo nombre', false);
  const nombreCompleto = [primerApellido, segundoApellido, primerNombre, segundoNombre].filter(Boolean).join(' ');

  return {
    ...acceso,
    CODIGO_VERIFICACION: codigo,
    PRIMER_APELLIDO: primerApellido,
    SEGUNDO_APELLIDO: segundoApellido,
    PRIMER_NOMBRE: primerNombre,
    SEGUNDO_NOMBRE: segundoNombre,
    NOMBRE_COMPLETO: nombreCompleto,
    EDAD: edad,
    CORREO_CONTACTO: email,
    TELEFONO: telefono,
    RESIDENCIA: residencia,
    OBSERVACIONES: limpiarTexto_(d.observaciones || '', 500),
    consentimiento: true
  };
}

function validarCampoNombre_(valor, etiqueta, obligatorio) {
  const s = String(valor || '').normalize('NFC').trim().replace(/\s+/g, ' ').toLocaleUpperCase('es-VE');
  if (!s && !obligatorio) return '';
  if (s.length < 2 || s.length > 60) throw new Error('El campo ' + etiqueta + ' debe contener entre 2 y 60 caracteres' + (obligatorio ? '.' : ' o dejarse vacío si no aplica.'));
  if (!/^[A-ZÁÉÍÓÚÜÑ]+(?:[ '\-][A-ZÁÉÍÓÚÜÑ]+)*$/.test(s)) {
    throw new Error('El campo ' + etiqueta + ' solo puede contener letras, espacios, apóstrofes o guiones. No use números.');
  }
  return s;
}

function registrarAuditoria_(ss, d, operacion, resultado) {
  const sh = ss.getSheetByName(CFG.AUDIT);
  if (!sh) return;
  sh.appendRow([new Date(), d.CEDULA, d.PERIODO, d.SUBPROYECTO, d.SECCION, operacion, resultado]);
}

function asegurarHoja_(ss, nombre, encabezados) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    sh.getRange(1, 1, 1, encabezados.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, encabezados.length);
  } else {
    const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    const map = mapaEncabezados_(current);
    const missing = encabezados.filter(h => map[h] === undefined);
    if (missing.length) {
      const start = sh.getLastColumn() + 1;
      sh.getRange(1, start, 1, missing.length).setValues([missing]);
      sh.getRange(1, start, 1, missing.length).setFontWeight('bold');
    }
  }
  return sh;
}

function mapaEncabezados_(encabezados) {
  const hm = {};
  encabezados.forEach((h, i) => {
    const key = String(h || '').trim().toUpperCase().replace(/\s+/g, '_');
    if (key) hm[key] = i;
  });
  return hm;
}

function normalizarCedulaTexto_(valor) {
  let s = String(valor === null || valor === undefined ? '' : valor).trim().toUpperCase().replace(/[\s.\-]/g, '');
  const m = s.match(/^([VE])(\d{5,9})$/);
  if (m) return m[1] + '-' + m[2];
  if (/^\d{5,9}$/.test(s)) return 'V-' + s;
  return '';
}

function normalizarTexto_(valor) {
  return String(valor === null || valor === undefined ? '' : valor)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[’‘]/g, "'")
    .trim().replace(/\s+/g, ' ').toUpperCase();
}

function limpiarTexto_(valor, maximo) {
  return String(valor === null || valor === undefined ? '' : valor).normalize('NFC').trim().replace(/\s+/g, ' ').slice(0, maximo);
}

function correoValido_(correo) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(correo || '').trim());
}

function enmascararCorreo_(correo) {
  const partes = String(correo).split('@');
  if (partes.length !== 2) return 'correo registrado';
  const local = partes[0];
  const visible = local.length <= 2 ? local.charAt(0) : local.slice(0, 2);
  return visible + '***@' + partes[1];
}

function hashTexto_(raw) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(raw), Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function claveOtp_(acceso) {
  const raw = [acceso.CEDULA, acceso.PERIODO, acceso.SUBPROYECTO, acceso.SECCION].join('|');
  return 'otp_' + hashTexto_(raw).slice(0, 40);
}
