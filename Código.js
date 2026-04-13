const SS = SpreadsheetApp.getActive();
const SH_CFG = SS.getSheetByName('Config');
const SH_EVT = SS.getSheetByName('Eventos');
const SH_PAR = SS.getSheetByName('Participantes');
const SH_ASI = SS.getSheetByName('Asistencias');
const SH_LIS = SS.getSheetByName('Listas');
const APP_VERSION = 'historial-v3-2025-10-14';

function pingVersion(){
  return { ok:true, v: APP_VERSION };
}

function setupSheets(){
  const ss = SpreadsheetApp.getActive();
  const must = [
    {name:'Eventos', headers:['id','nombre','fecha','max_participantes','estado','token','creado_en','actualizado_en']},
    {name:'Participantes', headers:['id','codigo','nombre','correo','telefono','creado_en']},
    {name:'Asistencias', headers:['id','event_id','participant_id','via','firma_file_id','firma_url','registrado_en']},
    {name:'Listas', headers:['id','event_id','total_participantes','archivo_tipo','archivo_file_id','archivo_url','generado_en']},
    {name:'Config', headers:['key','value']},
  ];
  must.forEach(s=>{
    let sh = ss.getSheetByName(s.name);
    if (!sh){
      sh = ss.insertSheet(s.name);
      sh.getRange(1,1,1,s.headers.length).setValues([s.headers]);
    } else {
      // Garantiza encabezados si la hoja está vacía
      if (sh.getLastRow() < 1) sh.getRange(1,1,1,s.headers.length).setValues([s.headers]);
    }
  });
  SpreadsheetApp.flush();
  return 'OK';
}

function cfg_(key){
  const r = SH_CFG.getRange(2,1,SH_CFG.getLastRow()-1,2).getValues();
  const f = r.find(row => String(row[0]).trim()==key);
  return f ? String(f[1]).trim() : '';
}
function now_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); }

function sh_(name){
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error("No existe la pestaña '"+name+"'");
  return sh;
}

function doGet(e) {
  const token = e?.parameter?.register || '';
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase();

  // 1) Con token => formulario de registro (interno; igual pedirá login del dominio)
  if (token) {
    const ev = getEventoByToken_(token);
    if (!ev) return HtmlService.createHtmlOutput('<h3>Evento no encontrado</h3>');
    if (ev.estado !== 'ABIERTO') return HtmlService.createHtmlOutput('<h3>Este evento está cerrado.</h3>');
    
    const t = HtmlService.createTemplateFromFile('Register');
    t.evento = ev;
    t.logoUrl = getLogoUrl();

    // 👇 Se añade meta viewport aquí para forzar responsive en móviles
    return t.evaluate()
      .setTitle('Registro de Asistencia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .addMetaTag('apple-mobile-web-app-capable', 'yes')
      .addMetaTag('mobile-web-app-capable', 'yes');
  }

  // 2) Sin token: si tu correo es admin => panel
  if (isAdminEmail_(me)) {
    const t = HtmlService.createTemplateFromFile('Admin');
    
    return t.evaluate()
      .setTitle('Asistencias – Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .addMetaTag('apple-mobile-web-app-capable', 'yes')
      .addMetaTag('mobile-web-app-capable', 'yes');
  }

  // 3) Sin token y no admin => denegado (solo gente del dominio llega aquí)
  return HtmlService.createHtmlOutput(
    '<h3>Acceso denegado</h3><p>No tienes permisos de administrador. Usa el enlace de registro del evento.</p>'
  );
}

function include_(fn){ return HtmlService.createHtmlOutputFromFile(fn).getContent(); }

// ====== EVENTOS ======

function getEventoById_(id){
  const vals = SH_EVT.getRange(2,1,SH_EVT.getLastRow()-1,8).getValues();
  const r = vals.find(v=>v[0]===id);
  if (!r) return null;
  return {id:r[0], nombre:r[1], fecha:r[2], max:r[3], estado:r[4], token:r[5]};
}
function getEventoByToken_(token){
  const vals = SH_EVT.getRange(2,1,SH_EVT.getLastRow()-1,8).getValues();
  const r = vals.find(v=>v[5]===token);
  if (!r) return null;
  return {id:r[0], nombre:r[1], fecha:r[2], max:r[3], estado:r[4], token:r[5]};
}

function createEvento(payload){
  try {
    const id = nextId_(SH_EVT,'EV',1);
    const token = Utilities.getUuid();
    const fecha = payload.fecha;
    const max = Number(payload.max || 0);
    const now = now_();
    SH_EVT.appendRow([id, payload.nombre, fecha, max, 'ABIERTO', token, now, now]);
    return JSON.stringify({ ok:true, evento:{ id, nombre:payload.nombre, fecha, max, estado:'ABIERTO', token } });
  } catch (err) {
    return JSON.stringify({ ok:false, error: String(err && err.message || err) });
  }
}


function closeEvento(eventId){
  const lr = SH_EVT.getLastRow();
  const r = SH_EVT.getRange(2,1,lr-1,8).getValues();
  for (let i=0;i<r.length;i++){
    if (r[i][0]===eventId){
      SH_EVT.getRange(i+2,5).setValue('CERRADO');
      SH_EVT.getRange(i+2,8).setValue(now_());
      return {ok:true};
    }
  }
  return {ok:false, msg:'Evento no encontrado'};
}

// ====== PARTICIPANTES & ASISTENCIAS ======

function registrarAsistencia({event_id, codigo, nombre, correo, telefono, firma_b64, via}){
  const ev = getEventoById_(event_id);
  if (!ev) return {ok:false, msg:'Evento inválido'};
  if (String(ev.estado).toUpperCase()!=='ABIERTO') return {ok:false, msg:'Evento cerrado'};

  const total = countAsistencias_(event_id);
  if (ev.max && total >= ev.max) return {ok:false, msg:'Cupo lleno'};

  const pid = upsertParticipante({codigo, nombre, correo, telefono});
  if (alreadyRegistered_(event_id, pid)){
    return {ok:false, code:'DUPLICATE', msg:'Esta persona ya está registrada en este evento'};
  }

  const {fileId, url} = saveFirma_(event_id, pid, firma_b64);

  const SH_ASI_ = sh_('Asistencias');
  const id = nextId_(SH_ASI_,'A',1);
  SH_ASI_.appendRow([id, event_id, pid, via||'public', fileId, url, now_()]);
  return {ok:true, id};
}

function upsertParticipante({codigo, nombre, correo, telefono}){
  const SH_PAR_ = sh_('Participantes');
  const lr = SH_PAR_.getLastRow();
  const vals = lr < 2 ? [] : SH_PAR_.getRange(2,1,lr-1,6).getValues(); // 6 cols
  let rowIndex = -1;

  if (codigo){
    rowIndex = vals.findIndex(v => String(v[1]).trim().toLowerCase() === String(codigo).trim().toLowerCase()); // B:codigo
  }
  if (rowIndex === -1 && correo){
    rowIndex = vals.findIndex(v => String(v[3]).trim().toLowerCase() === String(correo).trim().toLowerCase()); // D:correo
  }

  let id;
  if (rowIndex === -1){
    id = nextId_(SH_PAR_,'P',1);
    SH_PAR_.appendRow([id, codigo||'', nombre||'', correo||'', telefono||'', now_()]);
  } else {
    id = vals[rowIndex][0];
    const row = rowIndex+2;
    SH_PAR_.getRange(row,1,1,6).setValues([[id, codigo||vals[rowIndex][1], nombre||vals[rowIndex][2], correo||vals[rowIndex][3], telefono||vals[rowIndex][4], vals[rowIndex][5] ]]);
  }
  return id;
}

function countAsistencias_(eventId){
  const SH_ASI_ = sh_('Asistencias');
  const lr = SH_ASI_.getLastRow();
  if (lr < 2) return 0;
  const vals = SH_ASI_.getRange(2,1,lr-1,7).getValues();
  return vals.filter(v=>v[1]===eventId).length;
}

function saveFirma_(eventId, participantId, b64){
  const folderId = cfg_('FOLDER_FIRMAS_ID');
  if (!folderId) throw new Error('FOLDER_FIRMAS_ID no configurado en Config');
  const root = DriveApp.getFolderById(folderId);
  const sub = getOrCreateSub_(root, eventId);
  const bytes = Utilities.base64Decode(b64.split(',')[1]);
  const blob = Utilities.newBlob(bytes, 'image/png', `firma_${participantId}_${Date.now()}.png`);
  const f = sub.createFile(blob);
  return {fileId: f.getId(), url: 'https://drive.google.com/uc?export=view&id='+f.getId()};
}

function listAsistencias(eventId){
  const SH_ASI_ = sh_('Asistencias');
  const SH_PAR_ = sh_('Participantes');

  const lrA = SH_ASI_.getLastRow();
  if (lrA < 2) return [];

  const a = SH_ASI_.getRange(2,1,lrA-1,7).getValues().filter(v=>v[1]===eventId);

  const lrP = SH_PAR_.getLastRow();
  const p = lrP < 2 ? [] : SH_PAR_.getRange(2,1,lrP-1,6).getValues();
  const mapP = Object.fromEntries(p.map(r=>[r[0], {codigo:r[1], nombre:r[2], correo:r[3], telefono:r[4]}]));

  return a.map(r=>({
    asistencia_id:r[0], participant_id:r[2],
    codigo: mapP[r[2]]?.codigo||'',
    nombre: mapP[r[2]]?.nombre||'',
    correo: mapP[r[2]]?.correo||'',
    telefono: mapP[r[2]]?.telefono||'',
    firma_url:r[5], via:r[3], registrado_en:r[6]
  }));
}

// ====== LISTAS / REPORTES ======
function generarLista(eventId){
  try{
    const ev = getEventoById_(eventId);
    if (!ev) return JSON.stringify({ok:false, msg:'Evento inválido'});

    // — 1) Cerrar evento —
    const lr = SH_EVT.getLastRow();
    const rows = SH_EVT.getRange(2,1,lr-1,8).getValues();
    for (let i=0;i<rows.length;i++){
      if (rows[i][0]===eventId){
        SH_EVT.getRange(i+2,5).setValue('CERRADO');
        SH_EVT.getRange(i+2,8).setValue(now_());
        break;
      }
    }
    SpreadsheetApp.flush();

    // — 2) Hoja temporal —
    const data = listAsistencias(eventId);
    const tmp = SpreadsheetApp.create(`Lista ${ev.nombre} ${new Date().toISOString().slice(0,10)}`);
    const sh  = tmp.getActiveSheet();
    sh.clear();

    // ======= PARÁMETROS AJUSTABLES (con proporción automática) =======
    const LOGO_LONG = 180;   // Lado “largo” objetivo del logo (px)
    const SIG_LONG  = 50;   // Lado “largo” objetivo de cada firma (px)
    const SIG_PADDING = 3;   // Margen interno alrededor de la firma

    const COL_W_TEXT  = 180; // Ancho columnas texto (1..4)
    const COL_W_FIRMA = 300; // Ancho columna “Firma”
    // ================================================================

    // --- CABECERA: logo en A1; título en B1:E2 (evita solape) ---
    const meta = getLogoMeta_();
    const logoBlob = getLogoBlobForPdf_();
    if (logoBlob){
      const lw = Number(meta && meta.w || 0), lh = Number(meta && meta.h || 0);
      let tW = LOGO_LONG, tH = LOGO_LONG;
      if (lw && lh){
        const s = Math.min(LOGO_LONG / Math.max(lw, lh), 1);
        tW = Math.floor(lw * s);
        tH = Math.floor(lh * s);
      }
      placeImageCenteredInCell_(sh, logoBlob, 1, 1, tW, tH);
    }
    sh.getRange('B1:E1').merge();
    sh.getRange('B2:E2').merge();
    sh.getRange('B1').setValue(`Lista de Asistencia – ${ev.nombre}`)
      .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange('B2').setValue(
      `Fecha: ${new Date(ev.fecha).toLocaleDateString('es-DO',{weekday:'long', day:'2-digit', month:'long', year:'numeric'})}`
    ).setHorizontalAlignment('center').setFontStyle('italic');
    sh.insertRowsAfter(2, 1);

    // --- ENCABEZADOS (azul + blanco) ---
    sh.getRange(4,1,1,5).setValues([['Código','Nombre','Correo','Número telefónico','Firma']])
      .setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    sh.setFrozenRows(4);

    // Ancho columnas
    sh.setColumnWidths(1, 4, COL_W_TEXT);
    sh.setColumnWidth(5, COL_W_FIRMA);

    // --- 3) Escribe filas con HYPERLINK (para Excel) ---
    let r = 5;
    data.forEach(d=>{
      sh.getRange(r,1,1,4).setValues([[d.codigo||'', d.nombre||'', d.correo||'', d.telefono||'']]);
      centerRange_(sh.getRange(r,1,1,4));
      const url = String(d.firma_url||'').trim();
      if (url){
        sh.getRange(r,5).setFormula('=HYPERLINK("' + url.replace(/"/g,'""') + '","Ver firma")');
      } else {
        sh.getRange(r,5).setValue('');
      }
      if (r%2===1) sh.getRange(r,1,1,5).setBackground('#fafafa');
      r++;
    });

    // BORDES
    const tableRange = sh.getRange(4,1,Math.max(1, data.length)+1,5);
    tableRange.setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID);

    const repFolderId = cfg_('FOLDER_REPORTES_ID');
    if (!repFolderId) return JSON.stringify({ok:false, msg:'FOLDER_REPORTES_ID no configurado en Config'});
    const repFolder = DriveApp.getFolderById(repFolderId);

    const baseExport = 'https://docs.google.com/spreadsheets/d/'+tmp.getId()+'/export';
    const gid = sh.getSheetId();
    const commonQs = '&size=letter&sheetnames=false&printtitle=false&gridlines=false&fzr=true'
                   + '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5'
                   + '&horizontal_alignment=CENTER'
                   + '&fitw=true';

    // --- 4) Exportar XLSX (sin imágenes; con links) ---
    SpreadsheetApp.flush();
    const xlsxResp = UrlFetchApp.fetch(baseExport + '?format=xlsx&gid=' + gid, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    let xlsxFile = null;
    if (xlsxResp.getResponseCode() === 200){
      xlsxFile = repFolder.createFile(xlsxResp.getBlob().setName(`Lista_${ev.id}.xlsx`));
    }

    // --- 5) Reemplazar col 5 por IMÁGENES (para el PDF) — SIN getResizedBlob_ ---
    if (data.length) sh.getRange(5,5,data.length,1).clearContent();
    r = 5;
    data.forEach(d=>{
      const m = String(d.firma_url||'').match(/[\?&]id=([a-zA-Z0-9_-]+)/);
      const fileId = m ? m[1] : null;
      if (fileId){
        try{
          let raw = DriveApp.getFileById(fileId).getBlob();
          if (!/png|jpeg|jpg/i.test(raw.getContentType())) raw = raw.setContentType('image/png');

          // Calcula proporción por lado largo SIG_LONG
          let iw=0, ih=0;
          try{ const img = ImagesService.open(raw); iw = img.getWidth(); ih = img.getHeight(); }catch(e){}
          let targetW = SIG_LONG, targetH = SIG_LONG;
          if (iw && ih){
            const s = Math.min(SIG_LONG / Math.max(iw, ih), 1);
            targetW = Math.max(1, Math.floor(iw * s));
            targetH = Math.max(1, Math.floor(ih * s));
          }
          placeImageCenteredInCell_(sh, raw, r, 5, targetW, targetH, SIG_PADDING);

          // Asegura altura de fila suficiente
          const needH = targetH + SIG_PADDING*2;
          if (sh.getRowHeight(r) < needH) sh.setRowHeight(r, needH);
        }catch(e){
          centerRange_(sh.getRange(r,5,1,1));
        }
      } else {
        centerRange_(sh.getRange(r,5,1,1));
        const needH = SIG_LONG + SIG_PADDING*2;
        if (sh.getRowHeight(r) < needH) sh.setRowHeight(r, needH);
      }
      r++;
    });

    // re-bordes por si el insert tocó algo
    tableRange.setBorder(true,true,true,true,true,true,'#cbd5e1',SpreadsheetApp.BorderStyle.SOLID);

    // Asegura persistencia + render de imágenes
    SpreadsheetApp.flush();
    Utilities.sleep(900);

    // --- 6) Exportar PDF apaisado (gid + fit-to-width) ---
    const pdfResp = UrlFetchApp.fetch(baseExport + '?format=pdf&portrait=false&gid=' + gid + commonQs, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (pdfResp.getResponseCode() !== 200){
      return JSON.stringify({ok:false, msg:'No se pudo exportar a PDF ('+pdfResp.getResponseCode()+')'});
    }
    const pdfFile = repFolder.createFile(pdfResp.getBlob().setName(`Lista_${ev.id}.pdf`));

    // --- 7) Historial: PDF + XLSX ---
    const nowStr = now_();
    const idPdf = nextId_(SH_LIS,'L',1);
    SH_LIS.appendRow([idPdf, eventId, data.length, 'PDF',  pdfFile.getId(),  pdfFile.getUrl(),  nowStr]);
    if (xlsxFile){
      const idX = nextId_(SH_LIS,'L',1);
      SH_LIS.appendRow([idX, eventId, data.length, 'XLSX', xlsxFile.getId(), xlsxFile.getUrl(), nowStr]);
    }
    SpreadsheetApp.flush();

    // — 8) Limpieza y retorno —
    DriveApp.getFileById(tmp.getId()).setTrashed(true);

    const listaRecord = {
      lista_id: idPdf,
      event_id: eventId,
      evento_nombre: ev.nombre,
      evento_fecha: (ev.fecha instanceof Date)
        ? Utilities.formatDate(ev.fecha, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(ev.fecha||''),
      evento_estado: 'CERRADO',
      total: data.length,
      archivo_tipo: 'PDF',
      archivo_url: pdfFile.getUrl(),
      generado_en: nowStr
    };
    return JSON.stringify({ ok:true, pdfUrl: pdfFile.getUrl(), xlsxUrl: xlsxFile ? xlsxFile.getUrl() : '', lista: listaRecord });

  }catch(e){
    return JSON.stringify({ok:false, msg:String(e && e.message || e)});
  }
}

// ====== Utils ======
function getHistorialDesdeListas(){
  const lrL = SH_LIS.getLastRow();
  if (lrL < 2) return [];
  const rowsL = SH_LIS.getRange(2,1,lrL-1,7).getValues();
  const lrE = SH_EVT.getLastRow();
  const rowsE = lrE < 2 ? [] : SH_EVT.getRange(2,1,lrE-1,8).getValues();
  const mapE = Object.fromEntries(rowsE.map(r => [
    r[0], { nombre:r[1], fecha:r[2], estado:String(r[4]||'').trim().toUpperCase() }
  ]));

  const out = rowsL.map(r => {
    const ev = mapE[r[1]] || {};
    const fechaStr = (ev.fecha instanceof Date)
      ? Utilities.formatDate(ev.fecha, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(ev.fecha||'');
    return {
      lista_id: r[0],
      event_id: r[1],
      evento_nombre: ev.nombre || '(sin nombre)',
      evento_fecha: fechaStr,
      evento_estado: ev.estado || '',
      total: r[2],
      archivo_tipo: r[3],
      archivo_url: r[5],
      generado_en: String(r[6]||'')
    };
  }).sort((a,b) => String(b.generado_en).localeCompare(String(a.generado_en)));

  return out;
}

function getHistorialDesdeListasJSON(){
  try{
    const lrL = SH_LIS.getLastRow();
    let out = [];
    if (lrL >= 2){
      const rowsL = SH_LIS.getRange(2,1,lrL-1,7).getValues(); // [id,event_id,total,archivo_tipo,archivo_file_id,archivo_url,generado_en]
      const lrE = SH_EVT.getLastRow();
      const rowsE = lrE < 2 ? [] : SH_EVT.getRange(2,1,lrE-1,8).getValues(); // [id,nombre,fecha,max,estado,token,creado_en,actualizado_en]
      const mapE = Object.fromEntries(rowsE.map(r => [
        r[0], { nombre:r[1], fecha:r[2], estado:String(r[4]||'').trim().toUpperCase() }
      ]));
      out = rowsL.map(r => {
        const ev = mapE[r[1]] || {};
        return {
          lista_id: r[0],
          event_id: r[1],
          evento_nombre: ev.nombre || '(sin nombre)',
          // normaliza fecha a string para evitar serializaciones raras
          evento_fecha: (ev.fecha instanceof Date)
            ? Utilities.formatDate(ev.fecha, Session.getScriptTimeZone(), "yyyy-MM-dd")
            : String(ev.fecha||''),
          evento_estado: ev.estado || '',
          total: r[2],
          archivo_tipo: r[3],
          archivo_url: r[5],
          generado_en: String(r[6]||'')
        };
      }).sort((a,b) => String(b.generado_en).localeCompare(String(a.generado_en)));
    }
    console.log('[getHistorialDesdeListasJSON] items:', out.length);
    return JSON.stringify({ ok:true, data: out });
  }catch(e){
    console.error('[getHistorialDesdeListasJSON] ERR:', e);
    return JSON.stringify({ ok:false, error: String(e && e.message || e) });
  }
}

function debugCountListas(){
  const lr = SH_LIS.getLastRow();
  const arr = getHistorialDesdeListas();
  console.log('[debugCountListas] lr:', lr, 'items:', arr.length);
  return { lr, count: arr.length, sample: arr.slice(0,3) };
}


// Mantén esta por compatibilidad: reusa la canónica.
function getHistorialSoloCerrados(){
  return getHistorialDesdeListas();
}

/** Devuelve {id, w, h} del logo si es imagen (Drive API) */
function getLogoMeta_(){
  const id = getLogoFileId_();
  if (!id) return null;
  try{
    // Pide metadatos de imagen (si aplica)
    const meta = Drive.Files.get(id, { fields: 'id, mimeType, imageMediaMetadata(width,height)' });
    const w = meta && meta.imageMediaMetadata && meta.imageMediaMetadata.width;
    const h = meta && meta.imageMediaMetadata && meta.imageMediaMetadata.height;
    return { id: id, w: Number(w)||0, h: Number(h)||0, mime: meta && meta.mimeType || '' };
  }catch(e){
    return { id: id, w:0, h:0, mime:'' };
  }
}

/** Calcula tamaño objetivo manteniendo proporción dentro de maxW x maxH */
function fitSize_(srcW, srcH, maxW, maxH){
  if (!srcW || !srcH) return { w: Math.min(maxW, 200), h: Math.min(maxH, 60) };
  const scale = Math.min(maxW / srcW, maxH / srcH, 1);
  return { w: Math.floor(srcW * scale), h: Math.floor(srcH * scale) };
}

/** Obtiene el ID del archivo de logo, ya sea por Config.LOGO_FILE_ID o por getLogoUrl(). */
function getLogoFileId_(){
  let id = cfg_('LOGO_FILE_ID');
  if (!id){
    const u = getLogoUrl();
    const m = u && u.match(/id=([a-zA-Z0-9_-]+)/);
    id = m ? m[1] : '';
  }
  return id || '';
}

/**
 * Pide al Drive API el thumbnail del archivo (PNG/JPG) a un tamaño razonable.
 * size: pixels del lado “largo” (ej. 128, 256, 700…)
 * Devuelve Blob.
 */
function getDriveThumbnailBlob_(fileId, size){
  if (!fileId) return null;
  try{
    // Requiere Drive Advanced API habilitado.
    const file = Drive.Files.get(fileId, { fields: 'thumbnailLink' });
    let url = file && file.thumbnailLink;
    if (!url) return null;

    // thumbnailLink típicamente termina en "=s220". Ajustamos a nuestro tamaño.
    url = url.replace(/=s\d+$/, '=s' + Math.max(32, Math.min(size||256, 1600)));

    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return null;

    const blob = resp.getBlob();
    // Asegura un mime estándar
    const ct = blob.getContentType();
    if (!/png|jpeg|jpg/i.test(ct)) {
      return blob.setContentType('image/png');
    }
    return blob;
  }catch(e){
    return null;
  }
}

/** Data URL para el logo (rápido, con caché) – para formulario público. */
function getLogoDataUrl(size){
  const S = Math.max(48, Number(size)||128);
  const cache = CacheService.getScriptCache();
  const key = 'LOGO_DATA_URL_'+S;
  const cached = cache.get(key);
  if (cached) return cached;

  const id = getLogoFileId_();
  if (!id) return '';

  const blob = getDriveThumbnailBlob_(id, S*3); // un poco más grande para nitidez
  if (!blob) return '';

  const mime = blob.getContentType() || 'image/png';
  const b64 = Utilities.base64Encode(blob.getBytes());
  const dataUrl = 'data:'+mime+';base64,'+b64;
  cache.put(key, dataUrl, 6*60*60); // 6h
  return dataUrl;
}

/** Blob del logo para PDF – tamaño moderado para no romper límites. */
function getLogoBlobForPdf_(){
  const id = getLogoFileId_();
  if (!id) return null;
  // 700 px lado largo suele bastar para encajar en cabecera sin exceder 2 MB/1M px
  const blob = getDriveThumbnailBlob_(id, 700);
  return blob;
}

// redimensiona manteniendo proporción si excede maxW o maxH o >maxPixels o >maxBytes
function getResizedBlob_(blob, maxW, maxH, maxPixels, maxBytes){
  maxW = maxW || 800;
  maxH = maxH || 800;
  maxPixels = maxPixels || (1000*1000); // 1 millón
  maxBytes = maxBytes || (2*1024*1024); // 2 MB

  let img = ImagesService.open(blob);
  let w = img.getWidth(), h = img.getHeight();
  // si demasiados píxeles o excede dimensiones → escalar
  if (w*h > maxPixels || w > maxW || h > maxH){
    const scale = Math.min(maxW / w, maxH / h, Math.sqrt(maxPixels/(w*h)));
    w = Math.max(1, Math.floor(w * scale));
    h = Math.max(1, Math.floor(h * scale));
    img = ImagesService.open(blob).resize(w, h);
    blob = img.getBlob().setContentTypeFromExtension();
  }
  // si aún pesa mucho, baja un poco más
  if (blob.getBytes().length > maxBytes){
    const scale2 = 0.75; // otro 25% menos
    w = Math.max(1, Math.floor(w * scale2));
    h = Math.max(1, Math.floor(h * scale2));
    img = ImagesService.open(blob).resize(w, h);
    blob = img.getBlob().setContentTypeFromExtension();
  }
  return blob;
}

function placeImageCenteredInCell_(sheet, blob, row, col, imgW, imgH, padding){
  // Coloca una imagen en (row,col) centrada. Ajusta ancho de columna y alto de fila.
  padding = (padding == null ? 6 : padding);
  const range = sheet.getRange(row, col);

  // Asegura tamaño mínimo de la celda
  const colWidth = sheet.getColumnWidth(col);
  const needW = (imgW + padding*2);
  if (colWidth < needW) sheet.setColumnWidth(col, needW);

  const rowHeight = sheet.getRowHeight(row);
  const needH = (imgH + padding*2);
  if (rowHeight < needH) sheet.setRowHeight(row, needH);

  // Inserta y centra mediante offsets
  const img = sheet.insertImage(blob, col, row); // OverGridImage
  try{
    img.setWidth(imgW).setHeight(imgH);
    const newColWidth = sheet.getColumnWidth(col);
    const newRowHeight = sheet.getRowHeight(row);
    const offX = Math.max(0, Math.floor((newColWidth - imgW)/2));
    const offY = Math.max(0, Math.floor((newRowHeight - imgH)/2));
    img.setAnchorCell(range);
    img.setAnchorCellXOffset(offX);
    img.setAnchorCellYOffset(offY);
  }catch(e){}
  return img;
}

function centerRange_(range){
  // Centra texto horizontal y verticalmente
  range.setHorizontalAlignment('center');
  // Vertical Alignment en todo el bloque
  try{
    range.setVerticalAlignment('middle'); // en algunas cuentas Workspace está disponible
  }catch(e){
    // fallback: incrementa altura de fila; ya lo hacemos por firma
  }
}

function getOrCreateSub_(parent, name){
  const it = parent.getFoldersByName(name);
  return it.hasNext()? it.next() : parent.createFolder(name);
}
function nextId_(sheet, prefix, col){
  const lr = sheet.getLastRow();
  if (lr<2) return prefix+'00001';
  const last = sheet.getRange(lr,1).getValue(); // asume id en col 1
  const num = Number(String(last).replace(/\D/g,''))+1;
  return prefix + String(num).padStart(5,'0');
}

function isAdminEmail_(email){
  if (!email) return false;
  const allowed = (cfg_('ADMIN_EMAILS')||'')
      .split(';').map(s=>s.trim().toLowerCase()).filter(Boolean);
  return allowed.includes(String(email).toLowerCase());
}

function getRegisterUrl(eventId){
  const ev = getEventoById_(eventId);
  if (!ev) throw new Error('Evento inválido');

  // 1) Base pública si existe; si no, la del deployment actual
  let base = cfg_('PUBLIC_WEBAPP_URL') || ScriptApp.getService().getUrl();

  // Strip out /a/<domain>/ to support users with multiple accounts
  base = base.replace(/\/a\/macros\/[^/]+\//, '/macros/').replace(/\/a\/[^/]+\//, '/');

  // Asegura que es /exec real (algunos pegan /dev o sin /exec)
  if (!/\/exec(\?|$)/.test(base)) {
    // normaliza: quita query y fuerza /exec
    base = String(base).replace(/\/(dev|user|exec)?(\?.*)?$/, '') + '/exec';
  }

  // 2) authuser – evita el “choque de cuentas” en navegadores con varias sesiones
  // Toma la cuenta #0 si no estaba especificado (se puede ajustar a 1 si tu cuenta
  // primaria suele ser la del dominio y la #0 la personal).
  const hasAuthuser = /[?&]authuser=/.test(base);
  if (!hasAuthuser) {
    base += (base.indexOf('?') === -1 ? '?' : '&') + 'authuser=1';
  }

  // 3) agrega el token de registro
  const url = base + (base.indexOf('?') === -1 ? '?' : '&') + 'register=' + encodeURIComponent(ev.token);

  return url;
}


function apiListEventosActivos(){
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Eventos');
    if (!sh) return { ok:false, error:"No existe la pestaña 'Eventos'." };

    const lr = sh.getLastRow();
    if (lr < 2) return { ok:true, data: [] };

    // Asegura 8 columnas según tu esquema
    const values = sh.getRange(2, 1, lr - 1, 8).getValues();

    // Columna 5 = estado
    const data = values
      .filter(r => String(r[4]).trim().toUpperCase() === 'ABIERTO')
      .map(r => ({
        id: r[0],
        nombre: r[1],
        fecha: r[2],            // puede ser Date o string, el front sólo lo muestra
        max: r[3],
        estado: String(r[4]).trim().toUpperCase(),
        token: r[5]
      }));

    return { ok:true, data };
  } catch (err) {
    return { ok:false, error: String(err && err.message || err) };
  }
}

function apiListEventosActivosV2(){
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Eventos');
    if (!sh) return JSON.stringify({ ok:false, error:"No existe la pestaña 'Eventos'." });

    const lr = sh.getLastRow();
    if (lr < 2) return JSON.stringify({ ok:true, data: [] });

    // Asegura al menos 8 columnas A:H
    const values = sh.getRange(2, 1, lr - 1, 8).getValues();

    const data = values
      .filter(r => String(r[4]).trim().toUpperCase() === 'ABIERTO') // E: estado
      .map(r => ({
        id: r[0],
        nombre: r[1],
        fecha: r[2],
        max: r[3],
        estado: String(r[4]).trim().toUpperCase(),
        token: r[5]
      }));

    return JSON.stringify({ ok:true, data: data });
  } catch (err) {
    return JSON.stringify({ ok:false, error: String(err && err.message || err) });
  }
}

function apiListAsistenciasV1(eventId){
  try { return JSON.stringify({ok:true, data:listAsistencias(eventId)}); }
  catch(e){ return JSON.stringify({ok:false, error:String(e && e.message || e)}); }
}

function getLogoUrl(){
  // 1) Config prioritaria
  const cfgId = cfg_('LOGO_FILE_ID');
  if (cfgId){
    try {
      const f = DriveApp.getFileById(cfgId);
      return 'https://drive.google.com/uc?export=view&id=' + f.getId();
    } catch(e){}
  }

  // 2) Buscar dentro de carpetas "Assets"
  try {
    const assetsIt = DriveApp.searchFolders('title = "Assets" and trashed = false');
    while (assetsIt.hasNext()){
      const folder = assetsIt.next();
      const files = folder.searchFiles('title = "logo.png" and trashed = false');
      if (files.hasNext()){
        const f = files.next();
        return 'https://drive.google.com/uc?export=view&id=' + f.getId();
      }
    }
  } catch(e){}

  // 3) Último recurso: cualquier logo.png
  try {
    const any = DriveApp.searchFiles('title = "logo.png" and trashed = false');
    if (any.hasNext()){
      const f = any.next();
      return 'https://drive.google.com/uc?export=view&id=' + f.getId();
    }
  } catch(e){}

  return '';
}

function alreadyRegistered_(eventId, participantId){
  const SH_ASI_ = sh_('Asistencias');
  const lr = SH_ASI_.getLastRow();
  if (lr < 2) return false;
  const a = SH_ASI_.getRange(2,1,lr-1,7).getValues();
  return a.some(r => r[1]===eventId && r[2]===participantId);
}

function findParticipanteByCodigo(codigo){
  try{
    if (!codigo) return {ok:false, msg:'Código vacío'};
    const SH_PAR_ = sh_('Participantes');
    const lr = SH_PAR_.getLastRow();
    if (lr < 2) return {ok:false, msg:'No hay participantes'};
    const vals = SH_PAR_.getRange(2,1,lr-1,6).getValues(); // [id,codigo,nombre,correo,telefono,creado_en]
    const row = vals.find(v => String(v[1]).trim().toLowerCase() === String(codigo).trim().toLowerCase());
    if (!row) return {ok:false, msg:'No encontrado'};
    return {ok:true, participante:{ id:row[0], codigo:row[1], nombre:row[2], correo:row[3], telefono:row[4] }};
  }catch(e){
    return {ok:false, msg:String(e && e.message || e)};
  }
}

function testLogo(){
  const id = cfg_('LOGO_FILE_ID');
  if (!id) throw new Error('Config.LOGO_FILE_ID vacío');
  const f = DriveApp.getFileById(id);
  Logger.log('OK logo: %s (%s bytes)', f.getName(), f.getBlob().getBytes().length);
  Logger.log('URL: https://drive.google.com/uc?export=view&id=' + id);
}

function ping(){ return {ok:true, msg:'pong'}; }