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
    'EDAD', 'FECHA_NACIMIENTO', 'CORREO_CONTACTO', 'TELEFONO', 'RESIDENCIA', 'SUBPROYECTO',
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
        FECHA_NACIMIENTO: d.FECHA_NACIMIENTO,
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
  verificarClaveYVentana_(d);
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
  const mf = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(d.fechaNacimiento || ''));
  if (!mf) throw new Error('Indique su fecha de nacimiento.');
  const fnac = new Date(Number(mf[1]), Number(mf[2]) - 1, Number(mf[3]));
  const hoy = new Date();
  let edad = hoy.getFullYear() - fnac.getFullYear();
  if (hoy.getMonth() < fnac.getMonth() || (hoy.getMonth() === fnac.getMonth() && hoy.getDate() < fnac.getDate())) edad--;
  if (fnac.getMonth() !== Number(mf[2]) - 1 || edad < 15 || edad > 90) throw new Error('La fecha de nacimiento no es válida (se admiten edades de 15 a 90 años).');
  const email = String(d.correo || '').trim().toLowerCase();
  const emailConfirmacion = String(d.correoConfirmacion || '').trim().toLowerCase();
  if (!correoValido_(email)) throw new Error('Ingrese un correo electrónico válido.');
  const dominio = email.split('@')[1] || '';
  if (/^(gmial|gmai|gmil|gamil|gmail\.con|gmail\.co$|hotmial|hotmal|hotmail\.con|outlok|outlook\.con|yahoo\.con|yahooo|yaho\.)/.test(dominio)) {
    throw new Error('Revise el dominio de su correo (' + dominio + '): parece tener un error de escritura.');
  }
  if (email !== emailConfirmacion) throw new Error('El correo electrónico y su confirmación no coinciden.');
  const residencia = limpiarTexto_(d.residencia, 120);
  if (residencia.length < 3) throw new Error('Indique municipio y estado de residencia.');
  const telefono = normalizarTelefono_(d.telefono);
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
    FECHA_NACIMIENTO: fnac,
    CORREO_CONTACTO: email,
    TELEFONO: "'" + telefono,
    RESIDENCIA: protegerFormula_(residencia),
    OBSERVACIONES: protegerFormula_(limpiarTexto_(d.observaciones || '', 500)),
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
  if (/^(.)\1+$/.test(s.replace(/[ '\-]/g, '')) || /^(NO|NA|NINGUNO|NINGUNA|SIN|XX+|ASD+|ABC)$/.test(s)) {
    throw new Error('El campo ' + etiqueta + ' no parece un dato real. Escriba su nombre o apellido tal como aparece en su cédula.');
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

function normalizarTelefono_(v) {
  let s = String(v || '').replace(/[^\d+]/g, '').replace(/^(\+58|0058|58)/, '');
  if (/^4/.test(s)) s = '0' + s;
  if (!/^0(412|414|416|422|424|426)\d{7}$/.test(s)) {
    throw new Error('Ingrese un celular venezolano válido, por ejemplo 0414-1234567 (operadoras 0412, 0414, 0416, 0422, 0424, 0426).');
  }
  return s.slice(0, 4) + '-' + s.slice(4);
}

function protegerFormula_(t) { return /^[=+\-@]/.test(t) ? "'" + t : t; }

/** Menú en la hoja de cálculo para el docente. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Registro Académico')
    .addItem('1. Preparar sistema', 'prepararSistema')
    .addItem('2. Definir clave de clase', 'definirClaveClase')
    .addItem('3. Definir ventana de registro', 'definirVentana')
    .addItem('4. Reporte de pendientes', 'generarReportePendientes')
    .addItem('5. Resumen de avance', 'generarResumen')
    .addItem('6. Alertas de calidad de datos', 'revisarCalidad')
    .addToUi();
}

function definirClaveClase() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Clave de clase', 'Escriba la clave que dará a los estudiantes en clase (vacío = sin clave):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const p = PropertiesService.getScriptProperties();
  const c = r.getResponseText().trim();
  if (c) p.setProperty('CLAVE_CLASE', c); else p.deleteProperty('CLAVE_CLASE');
  ui.alert(c ? 'Clave guardada.' : 'Clave eliminada: el acceso queda sin clave.');
}

/** Lista quién está autorizado en Lista_Clases pero aún no se ha registrado. */
function generarReportePendientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ros = ss.getSheetByName(CFG.ROSTER).getDataRange().getDisplayValues();
  const reg = ss.getSheetByName(CFG.RECORDS).getDataRange().getDisplayValues();
  const hr = mapaEncabezados_(ros[0]), hg = mapaEncabezados_(reg[0]);
  const clave = (c, p, s) => [normalizarCedulaTexto_(c), String(p).trim().toUpperCase(), normalizarTexto_(s)].join('|');
  const hechos = new Set(reg.slice(1).map(f => clave(f[hg.CEDULA], f[hg.PERIODO], f[hg.SUBPROYECTO])));
  const filas = ros.slice(1).filter(f => f[hr.CEDULA] && !hechos.has(clave(f[hr.CEDULA], f[hr.PERIODO], f[hr.SUBPROYECTO])))
    .map(f => [f[hr.CEDULA], f[hr.PERIODO], f[hr.SUBPROYECTO], f[hr.SECCION]]);
  let sh = ss.getSheetByName('Pendientes') || ss.insertSheet('Pendientes');
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['CEDULA', 'PERIODO', 'SUBPROYECTO', 'SECCION']]).setFontWeight('bold');
  if (filas.length) sh.getRange(2, 1, filas.length, 4).setNumberFormat('@').setValues(filas);
  return filas.length;
}

function verificarClaveYVentana_(d) {
  const p = PropertiesService.getScriptProperties();
  const clave = p.getProperty('CLAVE_CLASE');
  if (clave && String(d.claveClase || '').trim().toUpperCase() !== clave.trim().toUpperCase()) {
    throw new Error('La clave de clase es incorrecta. Solicítela al docente.');
  }
  const ini = p.getProperty('VENTANA_INICIO'), fin = p.getProperty('VENTANA_FIN'), ahora = Date.now();
  const fmt = ms => Utilities.formatDate(new Date(Number(ms)), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  if (ini && ahora < Number(ini)) throw new Error('El registro aún no está abierto. Abre el ' + fmt(ini) + '.');
  if (fin && ahora > Number(fin)) throw new Error('El periodo de registro cerró el ' + fmt(fin) + '. Contacte al docente.');
}

/** El estudiante escribe solo su cédula y recibe las clases en las que está autorizado. */
function consultarAsignaciones(datos) {
  try {
    datos = datos || {};
    verificarClaveYVentana_(datos);
    const tipo = String(datos.tipoCedula || '').trim().toUpperCase();
    const num = String(datos.numeroCedula || '').replace(/\D/g, '');
    if (!['V', 'E'].includes(tipo) || !/^\d{5,9}$/.test(num)) throw new Error('Ingrese un tipo de cédula válido (V o E) y entre 5 y 9 dígitos, sin puntos ni guiones.');
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.ROSTER);
    if (!sh) throw new Error('El responsable académico aún no ha configurado la lista de clases.');
    const v = sh.getDataRange().getDisplayValues();
    const hm = mapaEncabezados_(v[0]);
    const items = [];
    v.slice(1).forEach(r => {
      if (normalizarCedulaTexto_(r[hm.CEDULA]) !== tipo + '-' + num) return;
      const sub = CFG.SUBPROYECTOS.find(s => normalizarTexto_(s) === normalizarTexto_(r[hm.SUBPROYECTO]));
      const per = String(r[hm.PERIODO]).trim().toUpperCase(), sec = String(r[hm.SECCION]).trim().toUpperCase();
      if (sub && CFG.PERIODOS.includes(per) && CFG.SECCIONES.includes(sec)) items.push({ periodo: per, subproyecto: sub, seccion: sec });
    });
    if (!items.length) return { ok: false, message: 'Esa cédula no aparece en la lista de clases. Verifique el tipo (V/E) y los dígitos, o contacte al docente.' };
    return { ok: true, items: items };
  } catch (e) {
    return { ok: false, message: e && e.message ? e.message : 'No se pudo consultar la lista de clases.' };
  }
}

function definirVentana() {
  const ui = SpreadsheetApp.getUi(), p = PropertiesService.getScriptProperties();
  const a = ui.prompt('Apertura', 'Fecha de apertura (AAAA-MM-DD). Vacío = sin límite:', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const b = ui.prompt('Cierre', 'Fecha de cierre, inclusive (AAAA-MM-DD). Vacío = sin límite:', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  const t1 = a.getResponseText().trim(), t2 = b.getResponseText().trim();
  const ok = t => !t || /^\d{4}-\d{2}-\d{2}$/.test(t);
  if (!ok(t1) || !ok(t2)) { ui.alert('Formato inválido. Use AAAA-MM-DD, por ejemplo 2026-11-02.'); return; }
  if (t1) p.setProperty('VENTANA_INICIO', String(new Date(t1 + 'T00:00:00').getTime())); else p.deleteProperty('VENTANA_INICIO');
  if (t2) p.setProperty('VENTANA_FIN', String(new Date(t2 + 'T23:59:59').getTime())); else p.deleteProperty('VENTANA_FIN');
  ui.alert('Ventana guardada. Se aplica a todos los accesos desde este momento.');
}

function volcar_(nombre, encabezado, filas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  sh.clear();
  sh.getRange(1, 1, 1, encabezado.length).setValues([encabezado]).setFontWeight('bold');
  if (filas.length) sh.getRange(2, 1, filas.length, encabezado.length).setNumberFormat('@').setValues(filas.map(f => f.map(String)));
  sh.setFrozenRows(1);
}

function generarResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ros = ss.getSheetByName(CFG.ROSTER).getDataRange().getDisplayValues();
  const reg = ss.getSheetByName(CFG.RECORDS).getDataRange().getValues();
  const hr = mapaEncabezados_(ros[0]), hg = mapaEncabezados_(reg[0]);
  const clave = (p, s, c) => [String(p).trim().toUpperCase(), normalizarTexto_(s), String(c).trim().toUpperCase()].join('|');
  const grupos = {};
  ros.slice(1).filter(r => r[hr.CEDULA]).forEach(r => {
    const k = clave(r[hr.PERIODO], r[hr.SUBPROYECTO], r[hr.SECCION]);
    if (!grupos[k]) grupos[k] = { p: String(r[hr.PERIODO]).trim().toUpperCase(), s: String(r[hr.SUBPROYECTO]).trim(), c: String(r[hr.SECCION]).trim().toUpperCase(), a: 0, r: 0, edades: [], m: 0 };
    grupos[k].a++;
  });
  reg.slice(1).forEach(f => {
    const g = grupos[clave(f[hg.PERIODO], f[hg.SUBPROYECTO], f[hg.SECCION])];
    if (!g) return;
    g.r++;
    if (Number(f[hg.EDAD])) g.edades.push(Number(f[hg.EDAD]));
    if (f[hg.ESTADO_CORREO] !== 'VERIFICADO') g.m++;
  });
  const filas = Object.keys(grupos).sort().map(k => {
    const g = grupos[k];
    const prom = g.edades.length ? (g.edades.reduce((x, y) => x + y, 0) / g.edades.length).toFixed(1) : '';
    return [g.p, g.s, g.c, g.a, g.r, g.a - g.r, (g.a ? Math.round(100 * g.r / g.a) : 0) + '%', prom, g.m];
  });
  volcar_('Resumen', ['PERIODO', 'SUBPROYECTO', 'SECCION', 'AUTORIZADOS', 'REGISTRADOS', 'PENDIENTES', 'AVANCE', 'EDAD_PROMEDIO', 'CORREOS_CON_CAMBIO_PENDIENTE'], filas);
  return filas.length;
}

function revisarCalidad() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const v = ss.getSheetByName(CFG.RECORDS).getDataRange().getDisplayValues();
  const h = mapaEncabezados_(v[0]), out = [], vistos = { C: {}, T: {} };
  v.slice(1).forEach(r => {
    const ci = r[h.CEDULA];
    [['C', String(r[h.CORREO_CONTACTO]).toLowerCase(), 'CORREO REPETIDO'], ['T', String(r[h.TELEFONO]).replace(/\D/g, ''), 'TELEFONO REPETIDO']].forEach(x => {
      if (!x[1]) return;
      if (vistos[x[0]][x[1]] && vistos[x[0]][x[1]] !== ci) out.push([x[2], ci, x[1] + ' (también en ' + vistos[x[0]][x[1]] + ')']);
      else vistos[x[0]][x[1]] = ci;
    });
    if (String(r[h.NOMBRE_COMPLETO]).trim().split(' ').length < 3) out.push(['NOMBRE CORTO', ci, r[h.NOMBRE_COMPLETO]]);
    if (r[h.ESTADO_CORREO] !== 'VERIFICADO') out.push(['CORREO SIN VERIFICAR', ci, r[h.CORREO_CONTACTO]]);
  });
  volcar_('Alertas', ['TIPO', 'CEDULA', 'DETALLE'], out);
  return out.length;
}
