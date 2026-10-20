/**
 * CARTERA DIOBEN — V2.3
 * Fuente principal: DF_MORA_ESTADOS > Comisión
 * Próxima cuota: Cartera Berex > último Originado + destination=commission
 * Resumen recaudo: Asignación_Estructurados > Como Vamo'
 *
 * Conserva por REFERENCIA:
 * - EXCLUIR CORREO
 * - STATUS
 * - COMENTARIO
 * - FECHA ÚLTIMA GESTIÓN
 *
 * NO ENVÍA CORREOS. Solo prepara Masivos_Correos.
 */

const CD2 = {
  TZ: 'America/Bogota',

  MORA_ID: '1jcPPhtF2YK3Kr7P_A0Mgh2OqhOfnVWB2to3UPoSH5tE',
  MORA_HOJA: 'Comisión',

  CARTERA_ID: '13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU',
  CARTERA_HOJA: 'Cartera',

  COMO_VAMO_ID: '15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM',
  COMO_VAMO_HOJA: "Como Vamo'",
  ASIGNACION_HOJA: 'Estructurado_Dioben',

  H: {
    RESUMEN: 'Resumen',
    SEGUIMIENTO: 'Seguimiento',
    HISTORICO: 'Histórico',
    MASIVOS: 'Masivos_Correos',
    CONFIG: 'Configuración'
  },

  TOLERANCIA: 2
};

const SEG2 = [
  'REFERENCIA',
  'MORA ACTUAL',
  'COMPORTAMIENTO',
  'PUNTUALIDAD HISTÓRICA',
  'DÍAS PROMEDIO ATRASO',
  'ÚLTIMO PAGO',
  'PRÓXIMO PAGO',
  'DÍAS PARA PAGO',
  'VALOR PRÓXIMO PAGO',
  'NIVEL DE ATENCIÓN',
  'EXCLUIR CORREO',
  'STATUS',
  'COMENTARIO',
  'FECHA ÚLTIMA GESTIÓN',
  'ENVIAR CORREO',
  'MOTIVO NO ENVÍO',
  'ACTUALIZADO',
  'FECHA_INICIO_MORA_1',
  'FECHA_PROMESA',
  'MORA_AL_CREAR_COMPROMISO'
];

const HIST2 = [
  'FECHA EVENTO',
  'REFERENCIA',
  'TIPO EVENTO',
  'VALOR ANTERIOR',
  'VALOR NUEVO',
  'MORA ACTUAL',
  'STATUS',
  'COMENTARIO',
  'ORIGEN'
];

const MAS2 = [
  'REFERENCIA',
  'MORA ACTUAL',
  'COMPORTAMIENTO',
  'PRÓXIMO PAGO',
  'DÍAS PARA PAGO',
  'VALOR PRÓXIMO PAGO',
  'STATUS',
  'TIPO CORREO',
  'LLAVE ENVÍO',
  'HABILITADO',
  'MOTIVO'
];



const STATUS21 = {
  POSITIVOS: [
    'Pago', 'Al día', 'Permanecido', 'Promesa de pago', 'Reestructuración',
    'Pago confirmado', 'Acuerdo de pago', 'Pago programado',
    'Débito programado', 'Caso solucionado'
  ],
  NEGATIVOS: [
    'No contesta', 'Ilocalizado', 'No quiere pagar', 'PaB incumplido',
    'Promesa incumplida', 'Reestructuración incumplida', 'Pago pendiente', 'Sin fondos', 'Volver a llamar',
    'Correo enviado / Sin respuesta', 'Número equivocado / sin contacto'
  ],
  NEUTROS: ['En gestión']
};

/* =========================================================
 * MENÚ
 * ========================================================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Cartera Dioben')
    .addItem('🚀 Instalar / actualizar V2.4', 'instalarCarteraDiobenV2')
    .addItem('🔄 Actualizar ahora', 'actualizarCarteraDiobenV2')
    .addItem('📊 Actualizar Resumen', 'actualizarResumenDiobenV2')
    .addItem('✉️ Actualizar Masivos', 'actualizarSalidaMasivosV2')
    .addSeparator()
    .addItem('⏰ Crear trigger diario 7 AM', 'crearTriggerDiarioV2')
    .addItem('🧹 Eliminar trigger diario', 'eliminarTriggerDiarioV2')
    .addToUi();
}


/* =========================================================
 * INSTALACIÓN
 * ========================================================= */

function instalarCarteraDiobenV2() {
  const ss = SpreadsheetApp.getActive();

  Object.values(CD2.H).forEach(nombre => {
    if (!ss.getSheetByName(nombre)) ss.insertSheet(nombre);
  });

  prepararEstructuraV2_(ss);
  actualizarCatalogoStatusV21_();
  actualizarCarteraDiobenV2();
}


function prepararEstructuraV2_(ss) {
  let sh = ss.getSheetByName(CD2.H.SEGUIMIENTO);
  if (!sh.getLastRow()) sh.getRange(1, 1, 1, SEG2.length).setValues([SEG2]);

  sh = ss.getSheetByName(CD2.H.HISTORICO);
  if (!sh.getLastRow()) sh.getRange(1, 1, 1, HIST2.length).setValues([HIST2]);

  sh = ss.getSheetByName(CD2.H.MASIVOS);
  if (!sh.getLastRow()) sh.getRange(1, 1, 1, MAS2.length).setValues([MAS2]);

  sh = ss.getSheetByName(CD2.H.CONFIG);
  if (!sh.getLastRow()) {
    sh.getRange(1, 1, 9, 3).setValues([
      ['PARÁMETRO', 'VALOR', 'DESCRIPCIÓN'],
      ['Moras incluidas', 'Al día, Mora 1', 'Universo activo'],
      ['Tolerancia cobro', 2, 'Días'],
      ['Actualización sugerida', '07:00', 'Trigger diario'],
      ['Fuente mora', 'DF_MORA_ESTADOS / Comisión', 'Estado e historial'],
      ['Fuente calendario', 'Cartera Berex / Cartera', 'Próxima comisión'],
      ['Fuente recaudo', "Asignación_Estructurados / Como Vamo'", 'Resumen Dioben'],
      ['Regla correo V2', 'Excluir o STATUS = no enviar', 'Todavía no envía correos'],
      ['Versión', 'V2', 'Cartera Dioben']
    ]);
  }
}


function actualizarCatalogoStatusV21_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CD2.H.CONFIG);
  if (!sh) return;

  sh.getRange('E1:G50').clearContent();
  sh.getRange('E1:G1').setValues([['STATUS', 'TIPO', 'CORREO']])
    .setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');

  const filas = []
    .concat(STATUS21.POSITIVOS.map(s => [s, 'POSITIVO', 'NO']))
    .concat(STATUS21.NEGATIVOS.map(s => [s, 'NEGATIVO', 'SÍ']))
    .concat(STATUS21.NEUTROS.map(s => [s, 'NEUTRO', 'NO']));

  sh.getRange(2, 5, filas.length, 3).setValues(filas);
}


/* =========================================================
 * ACTUALIZACIÓN PRINCIPAL
 * ========================================================= */

function actualizarCarteraDiobenV2() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('Ya existe una actualización en ejecución.');
  }

  try {
    const ss = SpreadsheetApp.getActive();

    Object.values(CD2.H).forEach(nombre => {
      if (!ss.getSheetByName(nombre)) ss.insertSheet(nombre);
    });

    prepararEstructuraV2_(ss);
    actualizarCatalogoStatusV21_();

    const shSeg = ss.getSheetByName(CD2.H.SEGUIMIENTO);
    const anterior = leerSeguimientoV2_(shSeg);

    const comision = leerComisionV2_();
    const clientes = construirClientesV2_(comision);

    // V2.3: asignación real + control de mora del mes anterior.
    const asignadas = leerReferenciasAsignadasV23_();
    const evaluacion = filtrarCarteraDiobenV23_(clientes, comision, asignadas);
    Object.keys(clientes).forEach(ref => {
      if (!evaluacion.permitidas.has(ref)) delete clientes[ref];
    });
    escribirValidacionV23_(ss, evaluacion.revision);

    // Solo necesitamos buscar calendario para referencias activas.
    const referencias = Object.keys(clientes);
    const calendarios = leerCalendarioCarteraV2_(referencias);

    Object.keys(clientes).forEach(ref => {
      const c = clientes[ref];
      const cal = calendarios[ref] || [];

      const prox = obtenerProximaCuotaV2_(
        cal,
        c.moraActual
      );

      c.proximoPago = prox.fecha;
      c.diasParaPago = prox.dias;
      c.valorProximoPago = prox.valor;

      c.nivelAtencion = nivelAtencionV2_(
        c.moraActual,
        c.comportamiento,
        c.diasParaPago
      );
    });

    registrarHistoricoV2_(ss, anterior, clientes);

    const ahora = new Date();

    const filas = Object.keys(clientes).map(ref => {
      const c = clientes[ref];
      const old = anterior[ref] || {};

      const excluir = old.excluir === true;
      let status = old.status || '';
      const comentario = old.comentario || '';

      let fechaGestion = old.fechaGestion || '';

      if (!fechaGestion && (String(status).trim() || String(comentario).trim())) {
        fechaGestion = ahora;
      }

      const fechaPromesa = fecha2_(old.fechaPromesa) || '';
      let moraCompromiso = old.moraCompromiso || '';
      if (fechaPromesa && !moraCompromiso) moraCompromiso = old.mora || c.moraActual;

      const vencimiento = evaluarCompromisoV24_(status, fechaPromesa, moraCompromiso, c.moraActual);
      if (vencimiento.incumplido) {
        const anteriorStatus = status;
        status = vencimiento.nuevoStatus;
        c.nivelAtencion = '🔴 Prioritario';
        registrarIncumplimientoV24_(ss, c.ref, c.moraActual, anteriorStatus, status, fechaPromesa, comentario);
      }

      const correo = evaluarCorreoV2_(excluir, status);

      // Fecha de primera detección del episodio actual de Mora 1.
      // Si ya estaba en Mora 1, nunca reiniciar la fecha.
      const esMora1 = text2_(c.moraActual) === 'MORA 1';
      const eraMora1 = text2_(old.mora) === 'MORA 1';
      const inicioMora1 = esMora1
        ? (eraMora1 && fecha2_(old.fechaInicioMora1)
            ? fecha2_(old.fechaInicioMora1)
            : ahora)
        : '';

      return [
        c.ref,
        c.moraActual,
        c.comportamiento,
        c.puntualidad,
        c.promedioAtraso,
        c.ultimoPago,
        c.proximoPago,
        c.diasParaPago,
        c.valorProximoPago,
        c.nivelAtencion,
        excluir,
        status,
        comentario,
        fechaGestion,
        correo.enviar,
        correo.motivo,
        ahora,
        inicioMora1,
        fechaPromesa,
        moraCompromiso
      ];
    });

    ordenarSeguimientoV2_(filas);
    escribirSeguimientoV2_(shSeg, filas);

    actualizarSalidaMasivosV2();
    actualizarResumenDiobenV2();

    SpreadsheetApp.getActive().toast(
      filas.length + ' clientes actualizados en V2.3',
      '✅ Cartera Dioben',
      5
    );

  } finally {
    lock.releaseLock();
  }
}


/* =========================================================
 * COMISIÓN
 * ========================================================= */


/**
 * V2.3: Lee referencias realmente asignadas a Dioben.
 * Ante error de lectura, aborta: nunca habilitar correos con asignación incierta.
 */
function leerReferenciasAsignadasV23_() {
  const sh = SpreadsheetApp.openById(CD2.COMO_VAMO_ID)
    .getSheetByName(CD2.ASIGNACION_HOJA);
  if (!sh) throw new Error('No existe Estructurado_Dioben. Se detiene por seguridad.');
  const last = sh.getLastRow();
  if (last < 2) throw new Error('Estructurado_Dioben no contiene referencias.');
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  const refs = new Set(vals.map(r => ref2_(r[0])).filter(Boolean));
  if (!refs.size) throw new Error('No hay referencias válidas en Estructurado_Dioben.');
  return refs;
}

/**
 * Compara el último estado de cada mes de Comisión por REFERENCIA.
 * Solo una transición Al día -> Mora 1 es Mora 1 nueva confirmada.
 * Las discrepancias con la hoja de asignación se dejan para revisión
 * y NO pasan a Masivos_Correos.
 */
function filtrarCarteraDiobenV23_(clientes, src, asignadas) {
  const hoy = new Date();
  const mesActual = Utilities.formatDate(hoy, CD2.TZ, 'yyyy-MM');
  const anterior = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  const mesAnterior = Utilities.formatDate(anterior, CD2.TZ, 'yyyy-MM');
  const porRef = {};

  src.filas.forEach((f, n) => {
    const ref = ref2_(f[src.idx.REF]);
    if (!ref || src.idx.FECHA < 0) return;
    const fecha = fecha2_(f[src.idx.FECHA]);
    if (!fecha) return;
    const mes = Utilities.formatDate(fecha, CD2.TZ, 'yyyy-MM');
    if (mes !== mesActual && mes !== mesAnterior) return;
    const mora = text2_(f[src.idx.MORA]);
    if (!mora) return;
    if (!porRef[ref]) porRef[ref] = {};
    const actual = porRef[ref][mes];
    const stamp = fecha.getTime();
    if (!actual || stamp > actual.stamp || (stamp === actual.stamp && n > actual.n)) {
      porRef[ref][mes] = {mora, stamp, n};
    }
  });

  const permitidas = new Set();
  const revision = [];
  Object.keys(clientes).forEach(ref => {
    const c = clientes[ref];
    const asignada = asignadas.has(ref);
    const datos = porRef[ref] || {};
    const prev = datos[mesAnterior] ? datos[mesAnterior].mora : '';
    const curr = datos[mesActual] ? datos[mesActual].mora : '';
    const moraActual = text2_(c.moraActual);

    let motivo = '';
    if (!asignada) motivo = 'NO ASIGNADA A DIOBEN';
    else if (!curr) motivo = 'SIN ESTADO DEL MES ACTUAL';
    else if (curr !== moraActual) motivo = 'MORA ACTUAL INCONSISTENTE';
    else if (moraActual === 'AL DIA') {
      permitidas.add(ref);
    } else if (moraActual === 'MORA 1') {
      if (prev === 'AL DIA') permitidas.add(ref);
      else if (!prev) motivo = 'MORA 1 SIN HISTORIAL DEL MES ANTERIOR';
      else motivo = 'MORA 1 NO NUEVA: MES ANTERIOR ' + prev;
    } else motivo = 'MORA NO PERMITIDA';

    if (motivo && asignada) revision.push([
      ref, prev || 'SIN REGISTRO', curr || 'SIN REGISTRO',
      motivo, new Date()
    ]);
  });
  return {permitidas, revision};
}

function escribirValidacionV23_(ss, filas) {
  const nombre = 'Validación Asignación';
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  sh.clearContents();
  const headers = ['REFERENCIA','MORA MES ANTERIOR','MORA MES ACTUAL','MOTIVO REVISIÓN','ACTUALIZADO'];
  sh.getRange(1,1,1,headers.length).setValues([headers])
    .setBackground('#1F4E78').setFontColor('#FFFFFF').setFontWeight('bold');
  if (filas.length) sh.getRange(2,1,filas.length,5).setValues(filas);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,5);
}

function leerComisionV2_() {
  const sh = SpreadsheetApp
    .openById(CD2.MORA_ID)
    .getSheetByName(CD2.MORA_HOJA);

  if (!sh) throw new Error('No existe la hoja Comisión.');

  const datos = sh.getDataRange().getValues();
  const mapa = {};

  datos[0].forEach((h, i) => mapa[head2_(h)] = i);

  const idx = {
    REF: col2_(mapa, ['REFERENCIA', 'REFERENCE']),
    FECHA: col2_(mapa, ['FECHA']),
    PAGO: col2_(mapa, ['PAGO']),
    ORIGEN: col2_(mapa, ['FECHA_ORIGEN']),
    COBRO: col2_(mapa, ['FECHA_COBRO']),
    FECHA_PAGO: col2_(mapa, ['FECHA_DE_PAGO']),
    MORA: col2_(mapa, ['MORA_STATUS'])
  };

  if (idx.REF < 0 || idx.MORA < 0) {
    throw new Error('Falta REFERENCIA o MORA_STATUS en Comisión.');
  }

  return {
    filas: datos.slice(1),
    idx
  };
}


function construirClientesV2_(src) {
  const { filas, idx } = src;
  const grupos = {};

  filas.forEach(fila => {
    const ref = ref2_(fila[idx.REF]);
    if (!ref) return;

    if (!grupos[ref]) grupos[ref] = [];
    grupos[ref].push(fila);
  });

  const salida = {};

  Object.keys(grupos).forEach(ref => {
    const filasRef = grupos[ref];

    const actual = filasRef
      .slice()
      .sort((a, b) => ordenFechaV2_(b, idx) - ordenFechaV2_(a, idx))
      .find(f => String(f[idx.MORA] || '').trim());

    if (!actual) return;

    const mora = mora2_(actual[idx.MORA]);
    const moraNorm = text2_(mora);

    if (!['AL DIA', 'MORA 1'].includes(moraNorm)) return;

    const comportamiento = comportamientoV2_(filasRef, idx);

    salida[ref] = {
      ref,
      moraActual: mora,
      comportamiento: comportamiento.clasificacion,
      puntualidad: comportamiento.puntualidad,
      promedioAtraso: comportamiento.promedioAtraso,
      ultimoPago: ultimoPagoV2_(filasRef, idx),
      proximoPago: '',
      diasParaPago: '',
      valorProximoPago: '',
      nivelAtencion: ''
    };
  });

  return salida;
}


function ordenFechaV2_(fila, idx) {
  for (const c of [idx.FECHA, idx.COBRO, idx.FECHA_PAGO, idx.ORIGEN]) {
    if (c < 0) continue;
    const d = fecha2_(fila[c]);
    if (d) return d.getTime();
  }
  return 0;
}


/* =========================================================
 * COMPORTAMIENTO — V2
 * =========================================================
 * Usa pagos históricos reales.
 * Un cliente sin suficientes pares FECHA_DE_PAGO + FECHA_COBRO
 * permanece "Sin historial"; no se castiga por ello.
 * ========================================================= */

function comportamientoV2_(filas, idx) {
  const hoy = dia2_(new Date());

  let total = 0;
  let puntuales = 0;
  let sumaAtrasos = 0;
  let atrasados = 0;

  filas.forEach(fila => {
    if (idx.FECHA_PAGO < 0 || idx.COBRO < 0) return;

    const esperado = fecha2_(fila[idx.FECHA_PAGO]);
    const cobro = fecha2_(fila[idx.COBRO]);

    if (!esperado || !cobro) return;
    if (dia2_(esperado) > hoy) return;

    const atraso = Math.max(
      0,
      diff2_(dia2_(esperado), dia2_(cobro))
    );

    total++;

    if (atraso <= CD2.TOLERANCIA) {
      puntuales++;
    } else {
      sumaAtrasos += atraso;
      atrasados++;
    }
  });

  if (!total) {
    return {
      clasificacion: '⚪ Sin historial',
      puntualidad: '',
      promedioAtraso: ''
    };
  }

  const puntualidad = puntuales / total;

  let clasificacion;

  if (puntualidad >= 0.95) clasificacion = '🟢 Excelente';
  else if (puntualidad >= 0.80) clasificacion = '🟢 Bueno';
  else if (puntualidad >= 0.60) clasificacion = '🟡 Irregular';
  else clasificacion = '🔴 Riesgoso';

  return {
    clasificacion,
    puntualidad,
    promedioAtraso:
      atrasados > 0
        ? Math.round((sumaAtrasos / atrasados) * 10) / 10
        : 0
  };
}


function ultimoPagoV2_(filas, idx) {
  if (idx.PAGO < 0 || idx.COBRO < 0) return '';

  let ultimo = null;

  filas.forEach(fila => {
    const pago = num2_(fila[idx.PAGO]);
    const cobro = fecha2_(fila[idx.COBRO]);

    if (pago !== null && pago > 0 && cobro) {
      if (!ultimo || cobro > ultimo) ultimo = cobro;
    }
  });

  return ultimo || '';
}


/* =========================================================
 * CARTERA BEREX — CALENDARIO VIGENTE
 * ========================================================= */

function leerCalendarioCarteraV2_(referencias) {
  if (!referencias.length) return {};

  const buscadas = new Set(referencias);

  const sh = SpreadsheetApp
    .openById(CD2.CARTERA_ID)
    .getSheetByName(CD2.CARTERA_HOJA);

  if (!sh) throw new Error('No existe la hoja Cartera.');

  const datos = sh.getDataRange().getValues();
  const mapa = {};

  datos[0].forEach((h, i) => mapa[head2_(h)] = i);

  const cRef = col2_(mapa, ['REFERENCE', 'REFERENCIA']);
  const cDest = col2_(mapa, ['DESTINATION']);
  const cFecha = col2_(mapa, ['PAYMENT_DATE']);
  const cAmount = col2_(mapa, ['AMOUNT']);
  const cNum = col2_(mapa, ['PAYMENT_NUMBER']);
  const cOrig = col2_(mapa, ['ORIGINADO']);

  if (cRef < 0 || cDest < 0 || cFecha < 0 || cOrig < 0) {
    throw new Error(
      'En Cartera faltan reference, destination, payment_date u Originado.'
    );
  }

  // 1. Último Originado por referencia.
  const ultimoOrigen = {};

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const ref = ref2_(fila[cRef]);

    if (!buscadas.has(ref)) continue;

    const origen = fecha2_(fila[cOrig]);
    if (!origen) continue;

    if (!ultimoOrigen[ref] || origen > ultimoOrigen[ref]) {
      ultimoOrigen[ref] = origen;
    }
  }

  // 2. Solo destination=commission de ese último Originado.
  const salida = {};

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const ref = ref2_(fila[cRef]);

    if (!buscadas.has(ref) || !ultimoOrigen[ref]) continue;
    if (text2_(fila[cDest]) !== 'COMMISSION') continue;

    const origen = fecha2_(fila[cOrig]);
    const fechaPago = fecha2_(fila[cFecha]);

    if (!origen || !fechaPago) continue;

    if (keyFecha2_(origen) !== keyFecha2_(ultimoOrigen[ref])) continue;

    if (!salida[ref]) salida[ref] = [];

    salida[ref].push({
      fecha: dia2_(fechaPago),
      valor: cAmount >= 0 ? num2_(fila[cAmount]) : '',
      numero: cNum >= 0 ? fila[cNum] : '',
      origen: dia2_(origen)
    });
  }

  Object.keys(salida).forEach(ref => {
    salida[ref].sort((a, b) => a.fecha - b.fecha);
  });

  return salida;
}


/* =========================================================
 * PRÓXIMA CUOTA — V2
 * =========================================================
 *
 * Si está Al día:
 *   toma primera cuota HOY o FUTURA.
 *   Nunca convierte una cuota vieja del calendario en mora.
 *
 * Si está Mora 1:
 *   intenta tomar la cuota más reciente vencida dentro del
 *   calendario vigente. Si no existe, toma la próxima.
 *
 * Esto evita el error antiguo de -190 / -1000 días.
 * ========================================================= */

function obtenerProximaCuotaV2_(calendario, moraActual) {
  if (!calendario.length) {
    return { fecha: '', dias: '', valor: '' };
  }

  const hoy = dia2_(new Date());
  const mora = text2_(moraActual);

  if (mora === 'AL DIA') {
    const cuota = calendario.find(c => c.fecha >= hoy);

    if (!cuota) {
      return { fecha: '', dias: '', valor: '' };
    }

    return {
      fecha: cuota.fecha,
      dias: diff2_(hoy, cuota.fecha),
      valor: cuota.valor
    };
  }

  if (mora === 'MORA 1') {
    const vencidas = calendario.filter(c => c.fecha < hoy);

    if (vencidas.length) {
      // La vencida más reciente, no la primera histórica.
      const cuota = vencidas[vencidas.length - 1];

      return {
        fecha: cuota.fecha,
        dias: diff2_(hoy, cuota.fecha),
        valor: cuota.valor
      };
    }

    const futura = calendario.find(c => c.fecha >= hoy);

    if (futura) {
      return {
        fecha: futura.fecha,
        dias: diff2_(hoy, futura.fecha),
        valor: futura.valor
      };
    }
  }

  return { fecha: '', dias: '', valor: '' };
}


/* =========================================================
 * NIVEL DE ATENCIÓN
 * ========================================================= */

function nivelAtencionV2_(mora, comportamiento, dias) {
  if (text2_(mora) === 'MORA 1') return '🔴 Hoy';

  const riesgoso = String(comportamiento).includes('Riesgoso');
  const irregular = String(comportamiento).includes('Irregular');

  if (typeof dias === 'number') {
    if (dias === 0) return '🔴 Hoy';

    if (dias === 1 && (riesgoso || irregular)) {
      return '🟠 Prioritario';
    }

    if (dias >= 2 && dias <= 3 && (riesgoso || irregular)) {
      return '🟡 Preventivo';
    }

    if (dias <= 2) return '🟡 Próximo';
  }

  if (riesgoso) return '🟠 Prioritario';
  if (irregular) return '🟡 Preventivo';

  return '🟢 Sin urgencia';
}


/* =========================================================
 * PERSISTENCIA MANUAL
 * ========================================================= */

function leerSeguimientoV2_(sh) {
  const out = {};

  if (sh.getLastRow() < 2) return out;

  const datos = sh.getDataRange().getValues();
  const mapa = {};

  datos[0].forEach((h, i) => mapa[head2_(h)] = i);

  const cRef = mapa[head2_('REFERENCIA')];

  if (cRef === undefined) return out;

  datos.slice(1).forEach(fila => {
    const ref = ref2_(fila[cRef]);
    if (!ref) return;

    out[ref] = {
      mora: val2_(fila, mapa, 'MORA ACTUAL'),
      excluir: val2_(fila, mapa, 'EXCLUIR CORREO') === true,
      status: val2_(fila, mapa, 'STATUS') || '',
      comentario: val2_(fila, mapa, 'COMENTARIO') || '',
      fechaGestion: val2_(fila, mapa, 'FECHA ÚLTIMA GESTIÓN') || '',
      fechaInicioMora1: val2_(fila, mapa, 'FECHA_INICIO_MORA_1') || '',
      fechaPromesa: val2_(fila, mapa, 'FECHA_PROMESA') || '',
      moraCompromiso: val2_(fila, mapa, 'MORA_AL_CREAR_COMPROMISO') || ''
    };
  });

  return out;
}


/* =========================================================
 * HISTÓRICO
 * ========================================================= */

function registrarHistoricoV2_(ss, anterior, clientes) {
  const sh = ss.getSheetByName(CD2.H.HISTORICO);
  const ahora = new Date();
  const eventos = [];

  Object.keys(clientes).forEach(ref => {
    const nuevo = clientes[ref];
    const viejo = anterior[ref];

    if (!viejo) {
      eventos.push([
        ahora, ref, 'ENTRADA A CARTERA', '',
        nuevo.moraActual, nuevo.moraActual,
        '', '', 'Actualización V2'
      ]);
      return;
    }

    if (text2_(viejo.mora) !== text2_(nuevo.moraActual)) {
      eventos.push([
        ahora, ref, 'CAMBIO DE MORA',
        viejo.mora, nuevo.moraActual,
        nuevo.moraActual,
        viejo.status || '',
        viejo.comentario || '',
        'Actualización V2'
      ]);
    }
  });

  Object.keys(anterior).forEach(ref => {
    if (clientes[ref]) return;

    eventos.push([
      ahora, ref, 'SALIDA DE CARTERA ACTIVA',
      anterior[ref].mora || '',
      'Fuera de Al día / Mora 1',
      '',
      anterior[ref].status || '',
      anterior[ref].comentario || '',
      'Actualización V2'
    ]);
  });

  if (eventos.length) {
    sh.getRange(
      sh.getLastRow() + 1,
      1,
      eventos.length,
      HIST2.length
    ).setValues(eventos);
  }

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HIST2.length).setFontWeight('bold');
}


/* =========================================================
 * CORREO
 * ========================================================= */

function evaluarCorreoV2_(excluir, status) {
  if (excluir === true) {
    return { enviar: 'NO', motivo: 'Exclusión manual' };
  }

  const tipo = clasificarStatusV21_(status);

  if (tipo === 'POSITIVO') {
    return { enviar: 'NO', motivo: 'Status positivo: ' + String(status || '').trim() };
  }

  if (tipo === 'NEUTRO') {
    return { enviar: 'NO', motivo: 'Status neutro / en gestión' };
  }

  if (tipo === 'NEGATIVO') {
    return { enviar: 'SÍ', motivo: 'Status negativo: ' + String(status || '').trim() };
  }

  return {
    enviar: 'SÍ',
    motivo: String(status || '').trim()
      ? 'Status no catalogado; revisar'
      : 'Sin status positivo'
  };
}

function clasificarStatusV21_(status) {
  const s = text2_(status);
  if (!s) return 'VACIO';
  if (STATUS21.POSITIVOS.some(x => text2_(x) === s)) return 'POSITIVO';
  if (STATUS21.NEGATIVOS.some(x => text2_(x) === s)) return 'NEGATIVO';
  if (STATUS21.NEUTROS.some(x => text2_(x) === s)) return 'NEUTRO';
  return 'NO_CATALOGADO';
}


function actualizarSalidaMasivosV2() {
  const ss = SpreadsheetApp.getActive();
  const seg = ss.getSheetByName(CD2.H.SEGUIMIENTO);
  const mas = ss.getSheetByName(CD2.H.MASIVOS);

  if (!seg || !mas) return;

  mas.clearContents();
  mas.getRange(1, 1, 1, MAS2.length).setValues([MAS2]);

  if (seg.getLastRow() < 2) return;

  const datos = seg.getDataRange().getValues();
  const mapa = {};

  datos[0].forEach((h, i) => mapa[head2_(h)] = i);

  const out = [];

  datos.slice(1).forEach(fila => {
    if (text2_(val2_(fila, mapa, 'ENVIAR CORREO')) !== 'SI') return;

    const ref = val2_(fila, mapa, 'REFERENCIA');
    const mora = val2_(fila, mapa, 'MORA ACTUAL');
    const prox = val2_(fila, mapa, 'PRÓXIMO PAGO');
    const dias = val2_(fila, mapa, 'DÍAS PARA PAGO');

    let tipo = 'SIN_REGLA';

    if (text2_(mora) === 'MORA 1') tipo = 'MORA_1';
    else if (dias === 0) tipo = 'PAGO_HOY';
    else if (dias === 2) tipo = 'PREVENTIVO_2_DIAS';
    else if (dias === 5) tipo = 'PREVENTIVO_5_DIAS';

    const fechaKey = fecha2_(prox)
      ? Utilities.formatDate(fecha2_(prox), CD2.TZ, 'yyyyMMdd')
      : 'SIN_FECHA';

    const llave = ref + '|' + fechaKey + '|' + tipo;

    out.push([
      ref,
      mora,
      val2_(fila, mapa, 'COMPORTAMIENTO'),
      prox,
      dias,
      val2_(fila, mapa, 'VALOR PRÓXIMO PAGO'),
      val2_(fila, mapa, 'STATUS'),
      tipo,
      llave,
      'SÍ',
      'Habilitado desde Seguimiento'
    ]);
  });

  if (out.length) {
    mas.getRange(2, 1, out.length, MAS2.length).setValues(out);
    mas.getRange(2, 4, out.length, 1).setNumberFormat('dd/mm/yyyy');
    mas.getRange(2, 6, out.length, 1).setNumberFormat('$#,##0');
  }

  mas.setFrozenRows(1);
  mas.getRange(1, 1, 1, MAS2.length)
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF');
}


/* =========================================================
 * ESCRITURA SEGUIMIENTO
 * ========================================================= */

function ordenarSeguimientoV2_(filas) {
  // Ordena filas completas (A:T), preservando checks, status, comentarios y fechas.
  const normalizar = valor => String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();

  function prioridad(fila) {
    const nivel = normalizar(fila[9]);    // J: NIVEL DE ATENCIÓN
    const status = normalizar(fila[11]);  // L: STATUS

    // Siempre primero las promesas o reestructuraciones incumplidas,
    // incluso si NIVEL DE ATENCIÓN sigue indicando otro valor.
    if (status.includes('PROMESA INCUMPLIDA') ||
        status.includes('REESTRUCTURACION INCUMPLIDA')) return 0;

    if (nivel.includes('HOY')) return 1;
    if (nivel.includes('PRIORITARIO')) return 2;
    if (nivel.includes('PROXIMO')) return 3;
    if (nivel.includes('PREVENTIVO')) return 4;
    if (nivel.includes('SIN URGENCIA')) return 5;
    return 99;
  }

  filas.sort((a, b) => {
    const pa = prioridad(a);
    const pb = prioridad(b);
    if (pa !== pb) return pa - pb;

    const da = typeof a[7] === 'number' ? a[7] : 99999;
    const db = typeof b[7] === 'number' ? b[7] : 99999;
    if (da !== db) return da - db;

    return String(a[0]).localeCompare(String(b[0]));
  });
}


function escribirSeguimientoV2_(sh, filas) {
  if (sh.getFilter()) sh.getFilter().remove();

  sh.clearContents();

  sh.getRange(1, 1, 1, SEG2.length).setValues([SEG2]);

  if (filas.length) {
    sh.getRange(2, 1, filas.length, SEG2.length).setValues(filas);
    sh.getRange(2, 11, filas.length, 1).insertCheckboxes();
    sh.getRange(1, 1, filas.length + 1, SEG2.length).createFilter();
  }

  formatoSeguimientoV2_(sh);
}


function formatoSeguimientoV2_(sh) {
  const n = sh.getLastRow();

  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  sh.getRange(1, 1, 1, SEG2.length)
    .setFontWeight('bold')
    .setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  if (n >= 2) {
    sh.getRange(2, 4, n - 1, 1).setNumberFormat('0.0%');

    // CORRECCIÓN V2: días promedio = número, no fecha.
    sh.getRange(2, 5, n - 1, 1).setNumberFormat('0.0');

    sh.getRange(2, 6, n - 1, 2).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 8, n - 1, 1).setNumberFormat('0');
    sh.getRange(2, 9, n - 1, 1).setNumberFormat('$#,##0');
    sh.getRange(2, 14, n - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange(2, 17, n - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sh.getRange(2, 18, n - 1, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(2, 19, n - 1, 1).setNumberFormat('dd/mm/yyyy').setBackground('#FFF2CC');

    // Zona manual.
    sh.getRange(2, 11, n - 1, 3).setBackground('#FFF2CC');

    const listaStatus = STATUS21.POSITIVOS.concat(STATUS21.NEGATIVOS, STATUS21.NEUTROS);
    const reglaStatus = SpreadsheetApp.newDataValidation()
      .requireValueInList(listaStatus, true)
      .setAllowInvalid(false)
      .build();

    const rangoStatus = sh.getRange(2, 12, n - 1, 1);
    rangoStatus.setDataValidation(reglaStatus);

    const reglasStatus = [];
    STATUS21.POSITIVOS.forEach(s => reglasStatus.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(s).setBackground('#D9EAD3').setFontColor('#274E13')
        .setRanges([rangoStatus]).build()
    ));
    STATUS21.NEGATIVOS.forEach(s => reglasStatus.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(s).setBackground('#F4CCCC').setFontColor('#990000')
        .setRanges([rangoStatus]).build()
    ));
    STATUS21.NEUTROS.forEach(s => reglasStatus.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(s).setBackground('#FFF2CC').setFontColor('#7F6000')
        .setRanges([rangoStatus]).build()
    ));
    sh.setConditionalFormatRules(reglasStatus);

  }

  const widths = [
    125, 110, 165, 150, 155, 125, 125, 115, 150,
    150, 140, 170, 380, 170, 120, 230, 170, 160, 145, 165
  ];

  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
}


/* =========================================================
 * RESUMEN V2
 * ========================================================= */

function actualizarResumenDiobenV2() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CD2.H.RESUMEN);
  const seg = ss.getSheetByName(CD2.H.SEGUIMIENTO);

  if (!sh || !seg) return;

  const r = resumenSeguimientoV2_(seg);
  const recaudo = leerComoVamoV2_();

  // Limpiar solo contenido/formato del tablero.
  sh.getRange('A1:H24').breakApart();
  sh.getRange('A1:H24').clear({ contentsOnly: false });

  // Título
  sh.getRange('A1:H1').merge()
    .setValue('CARTERA DIOBEN')
    .setFontSize(20)
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78')
    .setHorizontalAlignment('left');

  sh.getRange('A2:H2').merge()
    .setValue(
      'Resumen diario · ' +
      Utilities.formatDate(new Date(), CD2.TZ, 'dd/MM/yyyy HH:mm')
    )
    .setFontColor('#5F6368');

  // Tarjetas recaudo
  tarjetaV2_(sh, 'A4:B4', 'A5:B6', 'META MENSUAL', recaudo.meta, '$#,##0');
  tarjetaV2_(sh, 'C4:D4', 'C5:D6', 'RECAUDO', recaudo.lleva, '$#,##0');
  tarjetaV2_(sh, 'E4:F4', 'E5:F6', 'FALTA', recaudo.falta, '$#,##0');
  tarjetaV2_(sh, 'G4:H4', 'G5:H6', 'AVANCE', recaudo.avance, '0.0%');

  // Cartera
  sh.getRange('A8:H8').merge()
    .setValue('CARTERA ACTIVA')
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78');

  tarjetaV2_(sh, 'A9:B9', 'A10:B11', 'CLIENTES', r.total, '0');
  tarjetaV2_(sh, 'C9:D9', 'C10:D11', 'AL DÍA', r.alDia, '0');
  tarjetaV2_(sh, 'E9:F9', 'E10:F11', 'MORA 1', r.mora1, '0');
  tarjetaV2_(sh, 'G9:H9', 'G10:H11', 'ATENCIÓN HOY', r.hoy, '0');

  // Correos
  sh.getRange('A13:H13').merge()
    .setValue('GESTIÓN Y CORREOS')
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78');

  tarjetaV2_(sh, 'A14:B14', 'A15:B16', 'STATUS POSITIVO', r.statusPositivos, '0');
  tarjetaV2_(sh, 'C14:D14', 'C15:D16', 'STATUS NEGATIVO', r.statusNegativos, '0');
  tarjetaV2_(sh, 'E14:F14', 'E15:F16', 'HABILITADOS CORREO', r.habilitados, '0');
  tarjetaV2_(sh, 'G14:H14', 'G15:H16', 'EXCLUIDOS MANUAL', r.excluidos, '0');

  // Comportamiento
  sh.getRange('A18:H18').merge()
    .setValue('COMPORTAMIENTO DE PAGO')
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#1F4E78');

  sh.getRange('A19:H20').setValues([
    [
      '🟢 EXCELENTE', '', '🟢 BUENO', '',
      '🟡 IRREGULAR', '', '🔴 RIESGOSO', ''
    ],
    [
      r.excelente, '', r.bueno, '',
      r.irregular, '', r.riesgoso, ''
    ]
  ]);

  ['A19:B19','C19:D19','E19:F19','G19:H19',
   'A20:B20','C20:D20','E20:F20','G20:H20']
    .forEach(rango => sh.getRange(rango).merge());

  sh.getRange('A19:H19')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#D9EAF7');

  sh.getRange('A20:H20')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.setFrozenRows(2);

  for (let c = 1; c <= 8; c++) {
    sh.setColumnWidth(c, 135);
  }

  sh.setRowHeight(1, 34);
  sh.setRowHeights(5, 2, 30);
  sh.setRowHeights(10, 2, 30);
  sh.setRowHeights(15, 2, 30);
}


function tarjetaV2_(sh, rangoTitulo, rangoValor, titulo, valor, formato) {
  sh.getRange(rangoTitulo).merge()
    .setValue(titulo)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#D9EAF7')
    .setFontColor('#1F1F1F');

  sh.getRange(rangoValor).merge()
    .setValue(valor)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat(formato)
    .setBorder(true, true, true, true, false, false);
}


function resumenSeguimientoV2_(sh) {
  const r = {
    total: 0,
    alDia: 0,
    mora1: 0,
    hoy: 0,
    excluidos: 0,
    statusPositivos: 0,
    statusNegativos: 0,
    habilitados: 0,
    sinHistorial: 0,
    excelente: 0,
    bueno: 0,
    irregular: 0,
    riesgoso: 0
  };

  if (sh.getLastRow() < 2) return r;

  const datos = sh.getDataRange().getValues();
  const mapa = {};

  datos[0].forEach((h, i) => mapa[head2_(h)] = i);

  datos.slice(1).forEach(f => {
    const ref = val2_(f, mapa, 'REFERENCIA');
    if (!ref) return;

    r.total++;

    const mora = text2_(val2_(f, mapa, 'MORA ACTUAL'));
    const comp = String(val2_(f, mapa, 'COMPORTAMIENTO') || '');
    const nivel = String(val2_(f, mapa, 'NIVEL DE ATENCIÓN') || '');
    const status = String(val2_(f, mapa, 'STATUS') || '').trim();

    if (mora === 'AL DIA') r.alDia++;
    if (mora === 'MORA 1') r.mora1++;
    if (nivel.includes('🔴')) r.hoy++;

    if (val2_(f, mapa, 'EXCLUIR CORREO') === true) r.excluidos++;
    const tipoStatus = clasificarStatusV21_(status);
    if (tipoStatus === 'POSITIVO') r.statusPositivos++;
    if (tipoStatus === 'NEGATIVO') r.statusNegativos++;
    if (text2_(val2_(f, mapa, 'ENVIAR CORREO')) === 'SI') r.habilitados++;

    if (comp.includes('Sin historial')) r.sinHistorial++;
    else if (comp.includes('Excelente')) r.excelente++;
    else if (comp.includes('Bueno')) r.bueno++;
    else if (comp.includes('Irregular')) r.irregular++;
    else if (comp.includes('Riesgoso')) r.riesgoso++;
  });

  return r;
}


/* =========================================================
 * COMO VAMO' — LECTOR V2
 * =========================================================
 *
 * Primero intenta el bloque observado en tu archivo:
 *   DIOBEN alrededor de E20:H22
 *
 * Si cambia de posición, usa búsqueda flexible.
 * ========================================================= */

function leerComoVamoV2_() {
  const vacio = { meta: '', lleva: '', falta: '', avance: '' };

  try {
    const sh = SpreadsheetApp
      .openById(CD2.COMO_VAMO_ID)
      .getSheetByName(CD2.COMO_VAMO_HOJA);

    if (!sh) return vacio;

    // Según la estructura mostrada:
    // DIOBEN está en el bloque E:H, encabezados fila 21 y valores fila 22.
    const directo = sh.getRange('E20:H22').getValues();

    const textos = directo.flat().map(text2_);

    if (
      textos.includes('DIOBEN') &&
      textos.includes('META') &&
      textos.includes('LLEVA') &&
      textos.includes('FALTA') &&
      textos.includes('AVANCE')
    ) {
      return {
        meta: directo[2][0],
        lleva: directo[2][1],
        falta: directo[2][2],
        avance: directo[2][3]
      };
    }

    // Fallback flexible.
    const datos = sh.getDataRange().getValues();

    for (let r = 0; r < datos.length; r++) {
      for (let c = 0; c < datos[r].length; c++) {
        if (text2_(datos[r][c]) !== 'DIOBEN') continue;

        for (let rr = r + 1; rr <= Math.min(r + 4, datos.length - 2); rr++) {
          for (let cc = Math.max(0, c - 1); cc <= Math.min(c + 2, datos[rr].length - 4); cc++) {
            const h = [
              text2_(datos[rr][cc]),
              text2_(datos[rr][cc + 1]),
              text2_(datos[rr][cc + 2]),
              text2_(datos[rr][cc + 3])
            ];

            if (
              h[0] === 'META' &&
              h[1] === 'LLEVA' &&
              h[2] === 'FALTA' &&
              h[3] === 'AVANCE'
            ) {
              return {
                meta: datos[rr + 1][cc],
                lleva: datos[rr + 1][cc + 1],
                falta: datos[rr + 1][cc + 2],
                avance: datos[rr + 1][cc + 3]
              };
            }
          }
        }
      }
    }

    return vacio;

  } catch (e) {
    console.log('Error leyendo Como Vamo: ' + e.message);
    return vacio;
  }
}


/* =========================================================
 * TRIGGERS
 * ========================================================= */

function crearTriggerDiarioV2() {
  eliminarTriggerDiarioV2();

  ScriptApp.newTrigger('actualizarCarteraDiobenV2')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  SpreadsheetApp.getActive().toast(
    'Trigger V2 creado alrededor de las 7 AM',
    '⏰ Listo',
    5
  );
}


function eliminarTriggerDiarioV2() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'actualizarCarteraDiobenV2')
    .forEach(t => ScriptApp.deleteTrigger(t));
}


/* =========================================================
 * HELPERS
 * ========================================================= */

function head2_(v) {
  return String(v || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '_')
    .replace(/[^\w]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
}


function text2_(v) {
  return String(v || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}


function ref2_(v) {
  if (v === null || v === undefined || v === '') return '';

  return String(v).trim().replace(/\.0$/, '');
}


function mora2_(v) {
  const n = text2_(v);

  if (n === 'AL DIA') return 'Al día';
  if (n === 'MORA 1') return 'Mora 1';

  return String(v || '').trim();
}


function col2_(mapa, candidatos) {
  for (const c of candidatos) {
    const k = head2_(c);

    if (Object.prototype.hasOwnProperty.call(mapa, k)) {
      return mapa[k];
    }
  }

  return -1;
}


function val2_(fila, mapa, nombre) {
  const i = mapa[head2_(nombre)];
  return i === undefined ? '' : fila[i];
}


function fecha2_(v) {
  if (v === null || v === undefined || v === '') return null;

  if (
    Object.prototype.toString.call(v) === '[object Date]' &&
    !isNaN(v.getTime())
  ) {
    return new Date(v.getTime());
  }

  const d = new Date(v);

  return isNaN(d.getTime()) ? null : d;
}


function dia2_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}


function diff2_(desde, hasta) {
  return Math.round(
    (dia2_(hasta).getTime() - dia2_(desde).getTime()) / 86400000
  );
}


function keyFecha2_(d) {
  return Utilities.formatDate(d, CD2.TZ, 'yyyy-MM-dd');
}


function num2_(v) {
  if (v === null || v === undefined || v === '') return null;

  if (typeof v === 'number') {
    return isNaN(v) ? null : v;
  }

  const n = Number(
    String(v)
      .replace(/\$/g, '')
      .replace(/\s/g, '')
      .replace(/,/g, '')
  );

  return isNaN(n) ? null : n;
}


/* =========================================================
 * GESTIÓN MANUAL EN TIEMPO REAL — V2.1
 * ========================================================= */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CD2.H.SEGUIMIENTO) return;
    if (e.range.getRow() < 2 || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

    const col = e.range.getColumn();
    if (![11, 12, 13, 19].includes(col)) return;

    const row = e.range.getRow();
    const ref = ref2_(sh.getRange(row, 1).getValue());
    if (!ref) return;

    const mora = sh.getRange(row, 2).getValue();
    const excluir = sh.getRange(row, 11).getValue() === true;
    const status = sh.getRange(row, 12).getValue();
    const comentario = sh.getRange(row, 13).getValue();
    const ahora = new Date();
    const fechaPromesa = fecha2_(sh.getRange(row, 19).getValue());
    const tipoActual = text2_(status);
    if (col === 12 && ['PROMESA DE PAGO', 'REESTRUCTURACION'].includes(tipoActual)) {
      sh.getRange(row, 20).setValue(mora);
    }
    if (col === 19 && fechaPromesa && !sh.getRange(row, 20).getValue()) {
      sh.getRange(row, 20).setValue(mora);
    }

    sh.getRange(row, 14).setValue(ahora).setNumberFormat('dd/mm/yyyy hh:mm');

    const baseMora = sh.getRange(row, 20).getValue();
    const evaluacion = evaluarCompromisoV24_(status, fechaPromesa, baseMora, mora);
    let statusFinal = status;
    if (evaluacion.incumplido) {
      statusFinal = evaluacion.nuevoStatus;
      sh.getRange(row, 12).setValue(statusFinal);
      sh.getRange(row, 10).setValue('🔴 Prioritario');
    }
    const correo = evaluarCorreoV2_(excluir, statusFinal);
    sh.getRange(row, 15).setValue(correo.enviar);
    sh.getRange(row, 16).setValue(correo.motivo);

    const tipo = col === 11 ? 'CAMBIO EXCLUIR CORREO'
      : col === 12 ? 'CAMBIO DE STATUS'
      : col === 19 ? 'CAMBIO FECHA PROMESA'
      : 'CAMBIO DE COMENTARIO';

    const hist = e.source.getSheetByName(CD2.H.HISTORICO);
    if (hist) {
      hist.appendRow([
        ahora, ref, tipo,
        e.oldValue === undefined ? '' : e.oldValue,
        e.value === undefined ? '' : e.value,
        mora, statusFinal, comentario, 'Gestión manual V2.4'
      ]);
    }
  } catch (err) {
    console.log('onEdit V2.1: ' + err.message);
  }
}


/* =========================================================
 * COMPROMISOS — V2.4
 * FECHA_PROMESA: día comprometido; se vence a los 4 días calendario.
 * No se penaliza una mora que mejoró desde el compromiso.
 * ========================================================= */
function evaluarCompromisoV24_(status, fechaPromesa, moraInicial, moraActual) {
  const tipo = text2_(status);
  if (!['PROMESA DE PAGO', 'REESTRUCTURACION'].includes(tipo)) {
    return { incumplido: false };
  }
  const fecha = fecha2_(fechaPromesa);
  if (!fecha || !moraInicial) return { incumplido: false };

  const hoy = dia2_(new Date());
  const fechaLimite = dia2_(fecha);
  fechaLimite.setDate(fechaLimite.getDate() + 4);
  if (hoy < fechaLimite) return { incumplido: false };

  if (text2_(moraInicial) !== text2_(moraActual)) {
    return { incumplido: false };
  }

  return {
    incumplido: true,
    nuevoStatus: tipo === 'REESTRUCTURACION'
      ? 'Reestructuración incumplida'
      : 'Promesa incumplida'
  };
}

function registrarIncumplimientoV24_(ss, ref, mora, statusAnterior, statusNuevo, fecha, comentario) {
  const sh = ss.getSheetByName(CD2.H.HISTORICO);
  if (!sh) return;
  sh.appendRow([
    new Date(), ref, 'COMPROMISO INCUMPLIDO',
    statusAnterior, statusNuevo, mora, statusNuevo,
    'Fecha compromiso: ' + Utilities.formatDate(fecha, CD2.TZ, 'dd/MM/yyyy') +
      (comentario ? ' | ' + comentario : ''),
    'Automático V2.4'
  ]);
}