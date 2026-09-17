/* ============================================================
   Endpoint para la app de las cosmiatras.
   Recibe procedimientos terminados y los agrega a la pestana
   que le corresponde a cada persona.

   PRODUCTIVIDAD  (por defecto: las cosmiatras)
     A FECHA | B PACIENTE | C ATENCION | D ACTIVIDAD
     E ACTIVIDAD 2 | F ACTIVIDAD 3 | G INICIO | H FIN | I MINUTOS

   DOC ESTEFANY LOZA PEREZ  (mismas columnas; la C va vacia,
     igual que en todas las filas que esa hoja ya tenia)

   DOC GUSTAVO  (solo traia 3 columnas, asi que se le agregan
     INICIO, FIN y MINUTOS en D, E y F la primera vez)
     A Fecha | B Paciente | C Procedimiento | D INICIO | E FIN | F MINUTOS
   ============================================================ */

var TOKEN    = 'cosmiatras-2026';
var LIBRO_ID = '1gy2L1OpZewkXlhS93HB1n96PPtoJvm-QyKHqxUVi8r0';

/* A quien le toca cada pestana. Se compara el nombre en mayusculas
   buscando el trozo, asi da igual que en la app se llame
   "Dr. Gustavo", "GUSTAVO" o "Gustavo Perez". */
var DESTINOS = [
  { contiene: 'ESTEFANY', hoja: 'DOC ESTEFANY LOZA PEREZ', formato: 'estefany' },
  { contiene: 'GUSTAVO',  hoja: 'DOC GUSTAVO',             formato: 'gustavo'  }
];
var POR_DEFECTO = { contiene: '(las cosmiatras)', hoja: 'PRODUCTIVIDAD', formato: 'productividad' };

/* Lo que va en la columna ATENCION de la hoja de Estefany. Se escribe
   siempre esto, se llame como se llame en la app, para que las filas
   nuevas y las viejas queden iguales. */
var NOMBRE_ESTEFANY = 'DRA. ESTEFANY';

/* Donde cae cada dato en cada formato. cols = cuantas columnas se
   escriben de golpe; inicio/fin/minutos son numeros de columna. */
var FORMATOS = {
  productividad: { cols: 8, inicio: 7, fin: 8, minutos: 9 },
  estefany:      { cols: 8, inicio: 7, fin: 8, minutos: 9 },
  gustavo:       { cols: 5, inicio: 4, fin: 5, minutos: 6 }
};

function libro(){ return SpreadsheetApp.openById(LIBRO_ID); }

/* Busca la pestana ignorando espacios sobrantes: la de Estefany
   termina en espacio y es facil fallar por eso. */
function buscarHoja(nombre){
  var hojas = libro().getSheets();
  var busca = String(nombre).trim().toUpperCase();
  for (var i = 0; i < hojas.length; i++){
    if (hojas[i].getName().trim().toUpperCase() === busca) return hojas[i];
  }
  return null;
}

function destinoDe(persona){
  var p = String(persona || '').toUpperCase();
  for (var i = 0; i < DESTINOS.length; i++){
    if (p.indexOf(DESTINOS[i].contiene) >= 0) return DESTINOS[i];
  }
  return POR_DEFECTO;
}

/* Ultima fila con paciente (columna B en las tres hojas).
   No se usa getLastRow() porque las formulas de columnas enteras
   la falsean. */
function ultimaFila(hoja){
  var col = hoja.getRange(1, 2, hoja.getMaxRows(), 1).getValues();
  for (var i = col.length - 1; i >= 0; i--){
    if (String(col[i][0]).trim() !== '') return i + 1;
  }
  return 1;
}

/* Pestana oculta con los id ya escritos, para no duplicar si se
   reintenta un envio cuya respuesta se perdio. */
function hojaEnviados(){
  var lib = libro();
  var h = lib.getSheetByName('_enviados');
  if (!h){
    h = lib.insertSheet('_enviados');
    h.getRange('A1:B1').setValues([['id', 'cuando']]);
    h.hideSheet();
  }
  return h;
}

function yaEscritos(hojaEnv){
  var n = hojaEnv.getLastRow();
  var vistos = {};
  if (n < 2) return vistos;
  hojaEnv.getRange(2, 1, n - 1, 1).getValues().forEach(function(f){
    if (f[0]) vistos[String(f[0])] = true;
  });
  return vistos;
}

/* ---------- conversiones ---------- */
function aFecha(txt){                       /* '2026-09-15' -> numero de serie */
  var p = String(txt).split('-');
  /* Numero de serie y no un Date, para que la fecha no se corra un
     dia si la zona horaria del proyecto no es la de Lima. */
  return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000 + 25569;
}
function aHora(txt){                        /* '14:05' -> fraccion de dia */
  if (!txt) return '';
  var p = String(txt).split(':');
  return (+p[0] * 60 + +p[1]) / 1440;
}
function letra(n){                          /* 9 -> 'I' */
  var s = '';
  while (n > 0){
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* La hoja de Gustavo nacio con 3 columnas. Se le ponen los
   encabezados de hora la primera vez, sin tocar sus datos. */
function prepararGustavo(hoja){
  var faltan = 6 - hoja.getMaxColumns();
  if (faltan > 0) hoja.insertColumnsAfter(hoja.getMaxColumns(), faltan);
  if (String(hoja.getRange(1, 4).getValue()).trim() === ''){
    hoja.getRange(1, 4, 1, 3)
        .setValues([['INICIO DE ATENCION', 'FIN DE ATENCION', 'ATENCION MINUTOS']])
        .setFontWeight('bold');
  }
}

function filaSegun(formato, r){
  if (formato === 'gustavo'){
    return [aFecha(r.fecha), r.paciente, r.procedimiento, aHora(r.inicio), aHora(r.fin)];
  }
  if (formato === 'estefany'){
    return [aFecha(r.fecha), r.paciente, NOMBRE_ESTEFANY,
            r.procedimiento, '', '', aHora(r.inicio), aHora(r.fin)];
  }
  return [aFecha(r.fecha), r.paciente, String(r.cosmiatra).toUpperCase(),
          r.procedimiento, '', '', aHora(r.inicio), aHora(r.fin)];
}

function escribirEn(hoja, formato, registros){
  var conf = FORMATOS[formato];
  if (formato === 'gustavo') prepararGustavo(hoja);

  var fila    = ultimaFila(hoja) + 1;
  var valores = registros.map(function(r){ return filaSegun(formato, r); });

  var faltan = fila + valores.length - 1 - hoja.getMaxRows();
  if (faltan > 0) hoja.insertRowsAfter(hoja.getMaxRows(), faltan + 100);

  /* hereda el formato de la fila anterior para que se vea igual */
  if (fila > 2){
    hoja.getRange(fila - 1, 1, 1, conf.minutos)
        .copyTo(hoja.getRange(fila, 1, valores.length, conf.minutos), {formatOnly:true});
  }

  hoja.getRange(fila, 1, valores.length, conf.cols).setValues(valores);
  hoja.getRange(fila, 1, valores.length, 1).setNumberFormat('dd/mm/yyyy');
  hoja.getRange(fila, conf.inicio, valores.length, 2).setNumberFormat('h:mm');

  var ci = letra(conf.inicio), cf = letra(conf.fin);
  var formulas = [];
  for (var k = 0; k < valores.length; k++){
    var nf = fila + k;
    formulas.push(['=IF(AND(' + ci + nf + '<>"",' + cf + nf + '<>""),ROUND((' + cf + nf + '-' + ci + nf + ')*1440),"")']);
  }
  hoja.getRange(fila, conf.minutos, valores.length, 1).setFormulas(formulas);
  hoja.getRange(fila, conf.minutos, valores.length, 1).setNumberFormat('0');

  return fila;
}

/* ---------- entrada ---------- */
function doPost(e){
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (err) { return responder({ok:false, error:'ocupado, reintenta'}); }

  try {
    var datos = JSON.parse(e.postData.contents);
    if (datos.token !== TOKEN) return responder({ok:false, error:'token invalido'});

    var lista = datos.registros || [];
    if (!lista.length) return responder({ok:true, filas:0, ids:[]});

    var hojaEnv = hojaEnviados();
    var vistos  = yaEscritos(hojaEnv);

    /* se agrupan por pestana de destino, asi cada hoja se escribe
       de una sola vez aunque lleguen mezclados */
    var grupos = {}, orden = [], repetidos = [];
    lista.forEach(function(r){
      if (vistos[String(r.id)]){ repetidos.push(r.id); return; }
      vistos[String(r.id)] = true;
      var d = destinoDe(r.cosmiatra);
      if (!grupos[d.hoja]){ grupos[d.hoja] = { destino: d, regs: [] }; orden.push(d.hoja); }
      grupos[d.hoja].regs.push(r);
    });

    var detalle = [], escritos = [];
    orden.forEach(function(nombre){
      var g = grupos[nombre];
      var hoja = buscarHoja(g.destino.hoja);
      if (!hoja) throw new Error('No encuentro la pestana "' + g.destino.hoja + '"');

      var fila = escribirEn(hoja, g.destino.formato, g.regs);

      /* se apuntan los id en cuanto esa hoja quedo escrita, no al
         final: si algo falla despues, no se duplican al reintentar */
      var marca = new Date();
      hojaEnv.getRange(hojaEnv.getLastRow() + 1, 1, g.regs.length, 2)
             .setValues(g.regs.map(function(r){ return [r.id, marca]; }));

      g.regs.forEach(function(r){ escritos.push(r.id); });
      detalle.push({ hoja: hoja.getName(), filas: g.regs.length, desde: fila });
    });

    return responder({
      ok: true,
      filas: escritos.length,
      detalle: detalle,
      repetidos: repetidos,
      ids: escritos.concat(repetidos)
    });
  } catch (err) {
    return responder({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

/* Diagnostico. Sin token devuelve solo un saludo. */
function doGet(e){
  var tok = e && e.parameter ? e.parameter.token : '';
  if (tok !== TOKEN) return responder({ok:true, mensaje:'Endpoint activo'});
  try {
    var lib = libro();
    var salida = {
      ok: true,
      libro: lib.getName(),
      hojas: lib.getSheets().map(function(h){ return h.getName(); }),
      destinos: []
    };
    DESTINOS.concat([POR_DEFECTO]).forEach(function(d){
      var h = buscarHoja(d.hoja);
      salida.destinos.push({
        para:        d.contiene,
        hoja:        d.hoja,
        formato:     d.formato,
        encontrada:  !!h,
        columnas:    h ? h.getMaxColumns() : null,
        ultimaFila:  h ? ultimaFila(h) : null,
        proximaFila: h ? ultimaFila(h) + 1 : null
      });
    });
    return responder(salida);
  } catch (err) {
    return responder({ok:false, error:String(err)});
  }
}

function responder(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------------------------------------------------------
   Mantenimiento. Se ejecuta a mano desde el editor con el boton
   Ejecutar; no hace falta volver a desplegar para esto.
   No borra ni cambia datos: pone encabezados que faltaban, el
   nombre de Estefany donde estaba vacio, y formatos de vista.
   --------------------------------------------------------------- */
function arreglarHojas(){
  var informe = [];

  function formatos(hoja, etiqueta, colInicio, colMin){
    var n = Math.max(ultimaFila(hoja) - 1, 1);
    hoja.getRange(2, 1, n, 1).setNumberFormat('dd/mm/yyyy');
    hoja.getRange(2, colInicio, n, 2).setNumberFormat('h:mm');
    hoja.getRange(2, colMin, n, 1).setNumberFormat('0');
    if (hoja.getFrozenRows() < 1) hoja.setFrozenRows(1);
    informe.push(etiqueta + ': formatos aplicados a ' + n + ' filas');
  }

  /* --- DOC GUSTAVO: su hoja nacio con 3 columnas --- */
  var g = buscarHoja('DOC GUSTAVO');
  if (g){
    if (String(g.getRange(1, 4).getValue()).trim() === ''){
      g.getRange(1, 4, 1, 3)
       .setValues([['INICIO DE ATENCION', 'FIN DE ATENCION', 'ATENCION MINUTOS']]);
      informe.push('GUSTAVO: encabezados D, E y F puestos');
    } else {
      informe.push('GUSTAVO: los encabezados ya estaban');
    }
    /* que D, E y F se vean igual que A, B y C */
    g.getRange(1, 1, 1, 3).copyTo(g.getRange(1, 4, 1, 3), {formatOnly: true});
    g.setColumnWidth(4, 130);
    g.setColumnWidth(5, 130);
    g.setColumnWidth(6, 120);
    formatos(g, 'GUSTAVO', 4, 6);
  } else {
    informe.push('GUSTAVO: NO ENCUENTRO LA PESTANA');
  }

  /* --- DOC ESTEFANY: su nombre en la columna ATENCION --- */
  var e = buscarHoja('DOC ESTEFANY LOZA PEREZ');
  if (e){
    var ne = Math.max(ultimaFila(e) - 1, 1);
    var col = e.getRange(2, 3, ne, 1);
    var val = col.getValues();
    var puestos = 0;
    for (var i = 0; i < val.length; i++){
      if (String(val[i][0]).trim() === ''){ val[i][0] = NOMBRE_ESTEFANY; puestos++; }
    }
    if (puestos) col.setValues(val);
    informe.push('ESTEFANY: nombre puesto en ' + puestos + ' filas que lo tenian vacio');
    formatos(e, 'ESTEFANY', 7, 9);
  } else {
    informe.push('ESTEFANY: NO ENCUENTRO LA PESTANA');
  }

  /* --- PRODUCTIVIDAD: las fechas que se veian al reves --- */
  var p = buscarHoja('PRODUCTIVIDAD');
  if (p){
    var np = Math.max(ultimaFila(p) - 1, 1);
    p.getRange(2, 1, np, 1).setNumberFormat('dd/mm/yyyy');
    p.getRange(2, 7, np, 2).setNumberFormat('h:mm');
    if (p.getFrozenRows() < 1) p.setFrozenRows(1);
    informe.push('PRODUCTIVIDAD: fechas dd/mm/yyyy en ' + np + ' filas');
  }

  var texto = informe.join(' | ');
  Logger.log(texto);
  return texto;
}
