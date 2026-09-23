/**
 * AUTOCOM WAY · Mission Control
 * MVP 2026.09.17
 *
 * Línea base histórica: MC_Resultados
 * M01 cierre: 2026-08-18 · 500 pts base · recuperación 250
 * M02 cierre: 2026-08-31 · 500 pts base · recuperación 250
 * M03 cierre: 2026-09-24 · 500 pts · calidad SMA interna
 *
 * IMPORTANTE:
 * - MC_Resultados es la verdad histórica. No se recalculan M01/M02.
 * - Las recuperaciones y M03 se registran en hojas nuevas.
 * - El token NO se incluye en este archivo.
 */

const MC = {
 SHEET_ID: '1djdPT4dN17W7rqlJmCtUemaKZHQolMY7Rgij2-XeL6E',
 TZ: 'America/Mexico_City',

 BASE_RESULTS: 'MC_Resultados',
 MISSION_CATALOG: 'MC_Catalogo_Misiones',
 RECOVERY: 'MC_Recuperaciones',
 M01_SOURCE: 'M01_Sesion',
 M02_SOURCE: 'M02_Sesion',
 M02_REC_ROSTER: 'M02_Recuperacion_Padron',
 M02_REC_REG: 'M02_Recuperacion_Convocatoria',
 M02_REC_COVERAGE: 'M02_Recuperacion_Cobertura',
 M03_SAMPLE: 'M03_Muestreo',
 M03_CLOSE: 'M03_Cierre',

 MISSIONS: {
 M01: { id:'M01', name:'Activa tu unidad', base:500, recovery:250, deadline:'2026-08-18' },
 M02: { id:'M02', name:'Activa a tu equipo', base:500, recovery:250, deadline:'2026-08-31' },
 M03: { id:'M03', name:'Laboratorio y Muestreo de SMA', base:500, recovery:250, deadline:'2026-09-24' }
 },

 M02_RECOVERY: {
 capacity: 30,
 unlockAt: 27,
 sessions: [
 {id:'G1', label:'SMA Recuperación · Grupo 1', start:'2026-09-22T12:00:00', end:'2026-09-22T12:50:00', contingent:false},
 {id:'G2', label:'SMA Recuperación · Grupo 2', start:'2026-09-23T11:00:00', end:'2026-09-23T11:50:00', contingent:false},
 {id:'G3', label:'SMA Recuperación · Grupo 3', start:'2026-09-28T11:00:00', end:'2026-09-28T11:50:00', contingent:false}
 ]
 },

 SMA_CRITERIA: [
 {key:'control', label:'Mantiene el control del contacto'},
 {key:'situacion', label:'Parte de la situación u objeción real del prospecto'},
 {key:'concreta', label:'Define una actividad concreta y ejecutable'},
 {key:'fecha', label:'Tiene una fecha comprometida'},
 {key:'crm', label:'Puede registrarse y dar seguimiento en CRM'},
 {key:'feedback', label:'Genera feedback útil para avanzar'}
 ]
};

function doGet(e) {
 if (e && e.parameter && e.parameter.api === 'comites') {
 return responderComitesApi_(e);
 }
 return HtmlService.createTemplateFromFile('Index')
 .evaluate()
 .setTitle('AUTOCOM WAY · Mission Control')
 .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function responderComitesApi_(e) {
 const expected = PropertiesService.getScriptProperties().getProperty('MC_API_TOKEN') || '';
 const received = String((e.parameter && e.parameter.token) || '');
 if (!expected || received !== expected) {
 return jsonOutput_({ok:false, error:'No autorizado'});
 }
 return jsonOutput_({ok:true, data:getDirectorioComites_()});
}

function jsonOutput_(obj) {
 return ContentService.createTextOutput(JSON.stringify(obj))
 .setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
 return SpreadsheetApp.openById(MC.SHEET_ID);
}

function sheet_(name) {
 return ss_().getSheetByName(name);
}

function getOrCreateSheet_(name, headers) {
 let sh = sheet_(name);
 if (!sh) sh = ss_().insertSheet(name);
 if (sh.getLastRow() === 0 && headers && headers.length) {
 sh.getRange(1,1,1,headers.length).setValues([headers]);
 sh.setFrozenRows(1);
 }
 return sh;
}

function norm_(v) {
 return String(v == null ? '' : v)
 .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
 .trim().toUpperCase()
 .replace(/\s+/g,' ');
}

function bool_(v) {
 return v === true || ['SI','SÍ','TRUE','1','X'].indexOf(norm_(v)) >= 0;
}

function rowsAsObjects_(sh) {
 if (!sh || sh.getLastRow() < 2) return [];
 const values = sh.getDataRange().getValues();
 const headers = values.shift().map(String);
 return values.map(r => {
 const o = {};
 headers.forEach((h,i) => o[h] = r[i]);
 return o;
 });
}

function appendObject_(sh, headers, obj) {
 sh.appendRow(headers.map(h => obj[h] == null ? '' : obj[h]));
}

function now_() { return new Date(); }

function deadlineEnd_(dateStr) {
 return new Date(dateStr + 'T23:59:59-06:00');
}

function isExpired_(missionId) {
 return now_().getTime() > deadlineEnd_(MC.MISSIONS[missionId].deadline).getTime();
}

/* =========================
 BASE HISTÓRICA
 ========================= */

function getBaseRows_() {
 const sh = sheet_(MC.BASE_RESULTS);
 if (!sh) throw new Error('No existe la hoja MC_Resultados. Impórtala antes de publicar.');
 return rowsAsObjects_(sh);
}

function getUnits() {
 const rows = getBaseRows_();
 const map = {};
 rows.forEach(r => {
 const name = String(r.Unidad || '').trim();
 if (name) map[norm_(name)] = {id:String(r.Unidad_ID || ''), name:name};
 });
 return Object.keys(map).map(k => map[k]).sort((a,b)=>a.name.localeCompare(b.name,'es'));
}

function getHistorical_(unidad) {
 const key = norm_(unidad);
 const out = {M01:null, M02:null, points:0};
 getBaseRows_().forEach(r => {
 if (norm_(r.Unidad) !== key) return;
 const id = norm_(r.Mision_ID);
 if (id !== 'M01' && id !== 'M02') return;
 const item = {
 // Cumplida es el campo rector de la línea base rectificada.
 // Estado se conserva como dato descriptivo y NO puede cerrar una misión por sí solo.
 completed: norm_(r.Cumplida) === 'SI',
 points: Number(r.Puntos_total || 0),
 base: Number(r.Puntos_base || 0),
 bonus: Number(r.Bonus || 0),
 validation: String(r.Validacion || ''),
 updated: String(r.Fecha_actualizacion || '')
 };
 out[id] = item;
 out.points += item.points;
 });
 if (!out.M01) out.M01 = {completed:false,points:0,base:0,bonus:0};
 if (!out.M02) out.M02 = {completed:false,points:0,base:0,bonus:0};
 return out;
}

/* =========================
 RECUPERACIONES
 ========================= */

function recoveryHeaders_() {
 return ['Timestamp','Unidad','Mision_ID','Estado','Puntos_otorgados','Responsable','Correo','Detalle'];
}

function getRecoveryRows_(unidad) {
 const sh = sheet_(MC.RECOVERY);
 if (!sh) return [];
 const key = norm_(unidad);
 return rowsAsObjects_(sh).filter(r => norm_(r.Unidad) === key && norm_(r.Estado) === 'CUMPLIDA');
}

function getRecoveryMap_(unidad) {
 const map = {};
 getRecoveryRows_(unidad).forEach(r => {
 const id = norm_(r.Mision_ID);
 if (!map[id]) map[id] = {completed:true, points:Number(r.Puntos_otorgados || 0), timestamp:r.Timestamp};
 });
 return map;
}

function registerRecoveryM01(payload) {
 payload = payload || {};
 const unidad = String(payload.unidad || '').trim();
 if (!unidad) throw new Error('Selecciona una unidad.');

 const state = getMissionState(unidad);
 if (state.activeMission !== 'M01' || state.activeMode !== 'RECOVERY') {
 throw new Error('M01 no está habilitada para recuperación en esta unidad.');
 }

 const evidences = Array.isArray(payload.evidences) ? payload.evidences : [];
 if (evidences.length !== 5) throw new Error('La recuperación M01 requiere 5 evidencias.');

 const seenLinks = {};
 const seenProfiles = {};
 evidences.forEach((e,i) => {
 const link = String(e.link || '').trim();
 const profile = String(e.profile || '').trim();
 if (!link || !/^https?:\/\//i.test(link)) throw new Error('Revisa la liga del video ' + (i+1) + '.');
 if (!profile) throw new Error('Indica el perfil de la evidencia ' + (i+1) + '.');
 if (seenLinks[norm_(link)]) throw new Error('No repitas ligas de video.');
 if (seenProfiles[norm_(profile)]) throw new Error('Usa cinco perfiles diferentes.');
 seenLinks[norm_(link)] = true;
 seenProfiles[norm_(profile)] = true;
 });

 const lock = LockService.getScriptLock();
 lock.waitLock(15000);
 try {
 const recheck = getMissionState(unidad);
 if (recheck.activeMission !== 'M01') throw new Error('La misión ya cambió de estado.');

 const srcHeaders = ['Timestamp','Unidad','Mision','Nombre','Correo','Verificado','Liga del video','Participante','Comentario','Tipo'];
 const src = getOrCreateSheet_(MC.M01_SOURCE, srcHeaders);
 evidences.forEach(e => appendObject_(src, srcHeaders, {
 'Timestamp':now_(),'Unidad':unidad,'Mision':'M01',
 'Nombre':String(payload.responsable||''),'Correo':String(payload.correo||''),
 'Verificado':'RECUPERACION','Liga del video':String(e.link||''),
 'Participante':String(e.profile||''),'Comentario':String(e.comment||''),
 'Tipo':'RECUPERACION'
 }));

 completeRecovery_(unidad,'M01',payload.responsable,payload.correo,'5 evidencias de despliegue');
 return getMissionState(unidad);
 } finally {
 lock.releaseLock();
 }
}

/**
 * M02 Recuperación · convocatoria controlada.
 * - Padrón base: registros de M02_Cruce con CONECTÓ = No (95 filas fuente).
 * - 30 lugares por sesión.
 * - G3 visible como contingencia y se habilita cuando G1 y G2 llegan a 27.
 * - Confirmar convocatoria NO otorga puntos ni desbloquea M03.
 * - Los 250 pts se liberan sólo mediante validarM02RecoveryUnit_ después de validar asistencia.
 */
function m02RecoveryRosterSeed_() {
 return [{"unidad":"CHANGAN ZAMORA","email":"alberto.garcia@autocom.mx","nombre":""},{"unidad":"CORPORATIVO","email":"alejandra.martinezl@autocom.mx","nombre":"Alejandra Martínez"},{"unidad":"NISSAN ZITACUARO","email":"hilda.reynoso@autocom.mx","nombre":""},{"unidad":"NISSAN ZITACUARO","email":"benito.hernandez@autocom.mx","nombre":""},{"unidad":"TOYOTA REVOLUCION","email":"luciene.dasilva@autocom.mx","nombre":""},{"unidad":"TOYOTA REVOLUCION","email":"rocio.villegas@autocom.mx","nombre":""},{"unidad":"TOYOTA REVOLUCION","email":"roman.marin@autocom.mx","nombre":""},{"unidad":"TOYOTA REVOLUCION","email":"uriel.romero@autocom.mx","nombre":""},{"unidad":"CHANGAN MORELIA","email":"luz.moreno@autocom.mx","nombre":""},{"unidad":"CHANGAN MORELIA","email":"angel.iturbe@autocom.mx","nombre":""},{"unidad":"NISSAN URUAPAN","email":"hugo.paramo@autocom.mx","nombre":""},{"unidad":"NISSAN URUAPAN","email":"luis.perez@autocom.mx","nombre":""},{"unidad":"NISSAN URUAPAN","email":"marco.vargas@autocom.mx","nombre":""},{"unidad":"NISSAN URUAPAN","email":"raul.sandoval@autocom.mx","nombre":""},{"unidad":"NISSAN URUAPAN","email":"victor.aguilarz@autocom.mx","nombre":""},{"unidad":"NISSAN URUAPAN","email":"alvaro.alvarez@autocom.mx","nombre":""},{"unidad":"NISSAN URUAPAN","email":"alberto.soto@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"jorge.pena@autocom.mx","nombre":"ARMANDO PEÑA"},{"unidad":"TOYOTA UNIVERSIDAD","email":"cesar.monroy@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"jose.estrada@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"jose.juarez@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"jesus.mondragon@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"guillermo.enriquez@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"hugo.cruz@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"tania.escamilla@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"jacinto.martinez@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"nancy.monroy@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"uriel.hernandez@autocom.mx","nombre":""},{"unidad":"TOYOTA UNIVERSIDAD","email":"cesar.baeza@autocom.mx","nombre":""},{"unidad":"ACURA PEDREGAL","email":"cliserio.marin@autocom.mx","nombre":""},{"unidad":"ACURA PEDREGAL","email":"jose.guzman@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"america.cortes@autocom.mx","nombre":"América Mariana Cortes Vilchis"},{"unidad":"NISSAN JURIQUILLA","email":"ernesto.ponce@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"gabriela.gonzalez@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"javier.aguilar@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"jeronimo.balderas@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"marco.hernandez@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"mario.legorreta@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"paloma.rivera@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"rafael.carrera@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"teydi.rivera@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"angel.asencio@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"alfredo.guerrero@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"jorge.vilchez@autocom.mx","nombre":""},{"unidad":"NISSAN JURIQUILLA","email":"adriana.torres@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"america.cortes@autocom.mx","nombre":"América Mariana Cortes Vilchis"},{"unidad":"NISSAN LA CAPILLA","email":"ernesto.ponce@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"gabriela.gonzalez@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"javier.aguilar@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"jeronimo.balderas@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"marco.hernandez@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"mario.legorreta@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"america.cortes@autocom.mx","nombre":"América Mariana Cortes Vilchis"},{"unidad":"NISSAN LA CAPILLA","email":"paloma.rivera@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"rafael.carrera@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"teydi.rivera@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"angel.asencio@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"alfredo.guerrero@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"jorge.vilchez@autocom.mx","nombre":""},{"unidad":"NISSAN LA CAPILLA","email":"adriana.torres@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"diana.ocejo@autocom.mx","nombre":"Diana Patricia Ocejo Torres"},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"maricarmen.cardenas@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"isabel.cruz@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"victor.mejia@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"maria.padilla@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"brenda.chavero@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"paula.cardenas@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"alejandra.resendiz@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"javier.angeles@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"alejandro.barcenas@autocom.mx","nombre":""},{"unidad":"NISSAN SAN JUAN DEL RIO","email":"montserrat.gudino@autocom.mx","nombre":""},{"unidad":"NISSAN ZAMORA","email":"roberto.espinoza@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"josue.cardoso@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"oscar.perez@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"victor.vazquez@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"karen.galvan@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"lizbeth.serrano@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"eduardo.lopez@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"german.gastelum@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"antonio.hernandezo@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"ruben.olguin@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"julio.rodriguez@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"david.tovar@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"jaqueline.nieves@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"jose.ojeda@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"alfonso.gonzalez@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"froylan.santoyo@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"alejandro.ramos@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"ana.vilchis@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"yuliana.xolocotzi@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"erik.hernandez@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"omar.gutierrez@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"omar.reyes@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"domitila.trinidad@autocom.mx","nombre":""},{"unidad":"NISSAN BERNARDO QUINTANA","email":"gabino.quiroga@autocom.mx","nombre":""}];
}

function m02HistoricallyConnectedEmails_() {
 return new Set(["adan.ortega@autocom.mx", "alan.padilla@autocom.mx", "alberto.estrada@autocom.mx", "aldo.hernandez@autocom.mx", "aldo.ramirez@autocom.mx", "aleli.cruz@autocom.mx", "alfonso.nunez@autocom.mx", "alfonso.vaca@autocom.mx", "alondra.vargas@autocom.mx", "ana.duarte@autocom.mx", "angel.acosta@autocom.mx", "antonio.gonzalez@autocom.mx", "arcelia.negrete@autocom.mx", "ariana.rodriguez@autocom.mx", "aurora.gomez@autocom.mx", "axayacatl.lopez@autocom.mx", "blanca.alvez@autocom.mx", "brenda.mercado@autocom.mx", "bryan.perez@autocom.mx", "carlos.avila@autocom.mx", "cindy.plaza@autocom.mx", "daniel.delatorre@autocom.mx", "david.pavon@autocom.mx", "diana.roldan@autocom.mx", "diego.amezcua@autocom.mx", "dulce.ascencio@autocom.mx", "edgar.pardo@autocom.mx", "eduardo.rocha@autocom.mx", "edwin.miranda@autocom.mx", "elizabeth.tellez@autocom.mx", "emmanuel.gaytan@autocom.mx", "eric.soto@autocom.mx", "erik.villalobos@autocom.mx", "ernesto.herrera@autocom.mx", "fatima.vazquez@autocom.mx", "felipe.arambula@autocom.mx", "francisco.serrano@autocom.mx", "fred.ruiz@autocom.mx", "gabriela.navarro@autocom.mx", "gerardo.suarez@autocom.mx", "gloria.balderrama@autocom.mx", "guadalupe.martinez@autocom.mx", "guadalupe.trejo@autocom.mx", "horacio.martinez@autocom.mx", "hugo.santos@autocom.mx", "ismael.rodriguez@autocom.mx", "ivan.lua@autocom.mx", "ivan.mejia@autocom.mx", "ivan.santibanez@autocom.mx", "jaime.vilchis@autocom.mx", "janeth.guerrero@autocom.mx", "jenifer.estrella@autocom.mx", "jennifer.benitez@autocom.mx", "jesus.cervantes@autocom.mx", "jesus.ventura@autocom.mx", "jhonatan.monroy@autocom.mx", "jorge.alejandro@autocom.mx", "jorge.rios@autocom.mx", "jorge.santos@autocom.mx", "jorge.vazquez@autocom.mx", "jose.avila@autocom.mx", "jose.casanas@autocom.mx", "jose.hurtado@autocom.mx", "jose.mondragon@autocom.mx", "jose.nambo@autocom.mx", "jose.romero@autocom.mx", "jose.sanchez@autocom.mx", "jose.talavera@autocom.mx", "josemario.rodriguez@autocom.mx", "juan.ceron@autocom.mx", "juan.salvador@autocom.mx", "karen.leon@autocom.mx", "laura.luna@autocom.mx", "laura.solorio@autocom.mx", "leticia.avellaneda@autocom.mx", "lucio.sanchez@autocom.mx", "maira.castillo@autocom.mx", "marco.cruz@autocom.mx", "maria.carrillo@autocom.mx", "maria.falcon@autocom.mx", "maria.figueroa@autocom.mx", "maria.reyes@autocom.mx", "maria.valdemar@autocom.mx", "marial.sanchez@autocom.mx", "mario.acevedo@autocom.mx", "martin.ayala@autocom.mx", "mauricio.rivas@autocom.mx", "medardo.diaz@autocom.mx", "medren.arzate@autocom.mx", "miguel.maldonado@autocom.mx", "miguel.moreno@autocom.mx", "monica.lira@autocom.mx", "myrna.judith@autocom.mx", "norma.arevalo@autocom.mx", "octavio.rodriguez@autocom.mx", "omar.martinez@autocom.mx", "omar.olmedo@autocom.mx", "oscar.campillo@autocom.mx", "pedro.reyes@autocom.mx", "raymundo.perez@autocom.mx", "reynaldo.morales@autocom.mx", "ricardo.gutierrez@autocom.mx", "ricardo.sanchez@autocom.mx", "rodolfo.magana@autocom.mx", "salvador.navarro@autocom.mx", "serafin.riveros@autocom.mx", "silvia.barcena@autocom.mx", "soliris.algandar@autocom.mx", "uziel.albores@autocom.mx", "veronica.botello@autocom.mx", "viridiana.guevara@autocom.mx", "vladimir.santiago@autocom.mx"]);
}

function m02RosterHeaders_() { return ['Unidad','Correo','Nombre','Origen','Pendiente']; }
function m02RegHeaders_() { return ['Timestamp','Unidad','Correo_participante','Nombre_participante','Grupo_ID','Grupo','Inicio','Responsable','Correo_responsable','Estado','Asistencia']; }
function m02CoverageHeaders_() { return ['Unidad','Total_declarado','Responsable','Correo_responsable','Fecha_actualizacion']; }

function getM02Coverage_(unidad) {
 const sh=sheet_(MC.M02_REC_COVERAGE);
 const row=rowsAsObjects_(sh).find(r=>norm_(r.Unidad)===norm_(unidad));
 return row ? Math.max(0,Number(row.Total_declarado)||0) : 0;
}

function saveM02Coverage_(unidad,total,responsable,correo) {
 total=Math.floor(Number(total)||0);
 if (total < 1) throw new Error('Indica cuántos asesores deben participar en esta recuperación.');
 const sh=getOrCreateSheet_(MC.M02_REC_COVERAGE,m02CoverageHeaders_());
 const rows=rowsAsObjects_(sh);
 const idx=rows.findIndex(r=>norm_(r.Unidad)===norm_(unidad));
 const values=[unidad,total,String(responsable||''),String(correo||''),now_()];
 if (idx>=0) sh.getRange(idx+2,1,1,values.length).setValues([values]);
 else sh.appendRow(values);
 return total;
}

function configurarM02Recuperacion() {
 const rosterSh = getOrCreateSheet_(MC.M02_REC_ROSTER, m02RosterHeaders_());
 if (rosterSh.getLastRow() < 2) {
 const seed = m02RecoveryRosterSeed_();
 if (seed.length) rosterSh.getRange(2,1,seed.length,5).setValues(seed.map(x=>[x.unidad,x.email,x.nombre,'M02_Cruce · CONECTÓ=No','SI']));
 }
 getOrCreateSheet_(MC.M02_REC_REG, m02RegHeaders_());
 getOrCreateSheet_(MC.M02_REC_COVERAGE, m02CoverageHeaders_());

 const props = PropertiesService.getScriptProperties();
 MC.M02_RECOVERY.sessions.forEach(s => {
 const key = 'M02_REC_EVENT_' + s.id;
 let eventId = props.getProperty(key);
 let event = eventId ? CalendarApp.getDefaultCalendar().getEventById(eventId) : null;
 if (!event) {
 event = CalendarApp.getDefaultCalendar().createEvent(
 s.label,
 new Date(s.start + '-06:00'),
 new Date(s.end + '-06:00'),
 {description:'AUTOCOM WAY · Mission Control · M02 Recuperación'}
 );
 props.setProperty(key,event.getId());
 } else {
 // Mantiene el mismo evento e invitados, pero sincroniza nombre y horario con la configuración vigente.
 event.setTitle(s.label);
 event.setTime(new Date(s.start + '-06:00'), new Date(s.end + '-06:00'));
 }
 });
 return getM02RecoverySessions_();
}

function getM02RecoveryRegistrations_() {
 const sh = sheet_(MC.M02_REC_REG);
 return sh ? rowsAsObjects_(sh) : [];
}

function getM02RecoverySessions_() {
 const regs = getM02RecoveryRegistrations_();
 const counts = {G1:0,G2:0,G3:0};
 regs.forEach(r => { if (counts.hasOwnProperty(String(r.Grupo_ID))) counts[String(r.Grupo_ID)]++; });
 return MC.M02_RECOVERY.sessions.map(s => {
 const finalized = s.id === 'G1' || s.id === 'G2';
 const occupied = counts[s.id] || 0;
 return {
 id:s.id,label:s.label,start:s.start,end:s.end,capacity:MC.M02_RECOVERY.capacity,
 occupied:occupied,available:finalized ? 0 : Math.max(0,MC.M02_RECOVERY.capacity-occupied),
 contingent:!!s.contingent,finalized:finalized,unlocked:!finalized,
 full:finalized || occupied >= MC.M02_RECOVERY.capacity
 };
 });
}

function getM02RecoveryData_(unidad) {
 const key=norm_(unidad);
 const sh=sheet_(MC.M02_REC_ROSTER);
 const roster=sh ? rowsAsObjects_(sh).filter(r=>norm_(r.Unidad)===key && norm_(r.Pendiente)!=='NO') : [];
 const regs=getM02RecoveryRegistrations_();
 const byEmail={};
 regs.forEach(r=>{ byEmail[String(r.Correo_participante||'').trim().toLowerCase()]=r; });
 const people=roster.map(r=>{
 const email=String(r.Correo||'').trim().toLowerCase(), reg=byEmail[email]||null;
 return {email:email,name:String(r.Nombre||''),assigned:!!reg,groupId:reg?String(reg.Grupo_ID||''):'',group:reg?String(reg.Grupo||''):'',origin:String(r.Origen||'')};
 });
 const historicalRosterFound=roster.some(r=>norm_(r.Origen).indexOf('M02_CRUCE')>=0);
 const declaredTotal=getM02Coverage_(unidad);
 const assigned=people.filter(x=>x.assigned).length;
 return {people:people,sessions:getM02RecoverySessions_(),pending:people.filter(x=>!x.assigned).length,assigned:assigned,rosterFound:people.length>0,historicalRosterFound:historicalRosterFound,declaredTotal:declaredTotal,coverageRemaining:declaredTotal?Math.max(0,declaredTotal-assigned):null,coverageComplete:declaredTotal?assigned>=declaredTotal:false};
}

function registerRecoveryM02(payload) {
 throw new Error('La recuperación M02 ahora se realiza mediante convocatoria a sesiones. Actualiza la interfaz.');
}

function registerRecoveryM02Convocation(payload) {
 payload=payload||{};
 const unidad=String(payload.unidad||'').trim();
 const state=getMissionState(unidad);
 if (state.activeMission!=='M02' || state.activeMode!=='RECOVERY') throw new Error('M02 no está habilitada para recuperación en esta unidad.');

 const assignments=Array.isArray(payload.assignments)?payload.assignments:[];
 const manual=Array.isArray(payload.manualParticipants)?payload.manualParticipants:[];
 const declaredTotal=Math.floor(Number(payload.declaredTotal)||0);

 const lock=LockService.getScriptLock(); lock.waitLock(15000);
 try {
 const data=getM02RecoveryData_(unidad);
 // Sólo exigimos denominador declarado cuando NO existe padrón histórico de M02_Cruce.
 // Las capturas manuales previas no convierten una unidad en "padrón histórico conocido".
 let target=data.declaredTotal||0;
 if (!data.historicalRosterFound) {
 target=declaredTotal||target;
 if (target < 1) throw new Error('Indica cuántos asesores deben participar en esta recuperación.');
 const projected=data.assigned + assignments.length + manual.length;
 if (target < projected) throw new Error('El total declarado no puede ser menor que las personas convocadas.');
 saveM02Coverage_(unidad,target,payload.responsable,payload.correo);
 }
 if (!assignments.length && !manual.length) return getMissionState(unidad);

 const roster={}; data.people.forEach(p=>roster[p.email]={email:p.email,name:p.name,assigned:p.assigned,origin:p.origin||'M02_CRUCE'});
 const sessions={}; getM02RecoverySessions_().forEach(s=>sessions[s.id]=s);
 const connected=m02HistoricallyConnectedEmails_();
 const seen={};
 const all=[];

 assignments.forEach(a=>{
 const email=String(a.email||'').trim().toLowerCase(), gid=String(a.groupId||'');
 if (!roster[email]) throw new Error('El correo '+email+' no pertenece al padrón pendiente de esta unidad.');
 if (roster[email].assigned) throw new Error(email+' ya tiene una sesión asignada.');
 all.push({email:email,name:roster[email].name,groupId:gid,origin:'M02_CRUCE'});
 });

 manual.forEach(p=>{
 const name=String(p.name||'').trim(), email=String(p.email||'').trim().toLowerCase(), gid=String(p.groupId||'');
 if (!name) throw new Error('Captura el nombre de cada participante agregado.');
 if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Correo no válido: '+email);
 if (connected.has(email)) throw new Error(email+' ya aparece históricamente como conectado en M02; no requiere recuperación.');
 if (roster[email] && roster[email].assigned) throw new Error(email+' ya tiene una sesión asignada.');
 all.push({email:email,name:name,groupId:gid,origin:'CAPTURA_RESPONSABLE'});
 });

 all.forEach(p=>{
 if (seen[p.email]) throw new Error(p.email+' está repetido en la convocatoria.');
 if (!sessions[p.groupId]) throw new Error('Selecciona una sesión válida para '+p.email+'.');
 if (!sessions[p.groupId].unlocked) throw new Error('El Grupo 3 aún está bloqueado.');
 if (sessions[p.groupId].full) throw new Error('La sesión '+p.groupId+' ya está llena.');
 seen[p.email]=p.groupId;
 });

 const addCount={G1:0,G2:0,G3:0}; Object.values(seen).forEach(g=>addCount[g]++);
 Object.keys(addCount).forEach(g=>{ if (sessions[g] && sessions[g].occupied+addCount[g] > sessions[g].capacity) throw new Error('La sesión '+g+' no tiene lugares suficientes.'); });

 // Primero persiste los participantes que el responsable tuvo que complementar.
 const rosterSh=getOrCreateSheet_(MC.M02_REC_ROSTER,m02RosterHeaders_());
 manual.forEach(p=>{
 const email=String(p.email||'').trim().toLowerCase();
 if (!roster[email]) {
 appendObject_(rosterSh,m02RosterHeaders_(),{'Unidad':unidad,'Correo':email,'Nombre':String(p.name||'').trim(),'Origen':'CAPTURA_RESPONSABLE','Pendiente':'SI'});
 roster[email]={email:email,name:String(p.name||'').trim(),assigned:false,origin:'CAPTURA_RESPONSABLE'};
 }
 });

 const props=PropertiesService.getScriptProperties();
 const sh=getOrCreateSheet_(MC.M02_REC_REG,m02RegHeaders_());
 all.forEach(p=>{
 const s=sessions[p.groupId];
 const eventId=props.getProperty('M02_REC_EVENT_'+p.groupId);
 const event=eventId ? CalendarApp.getDefaultCalendar().getEventById(eventId) : null;
 if (!event) throw new Error('No encuentro el evento de '+p.groupId+'. Ejecuta configurarM02Recuperacion().');
 event.addGuest(p.email);
 appendObject_(sh,m02RegHeaders_(),{
 'Timestamp':now_(),'Unidad':unidad,'Correo_participante':p.email,'Nombre_participante':p.name,
 'Grupo_ID':p.groupId,'Grupo':s.label,'Inicio':new Date(s.start+'-06:00'),
 'Responsable':String(payload.responsable||''),'Correo_responsable':String(payload.correo||''),
 'Estado':'CONVOCADO','Asistencia':'PENDIENTE'
 });
 });
 return getMissionState(unidad);
 } finally { lock.releaseLock(); }
}

/** Puerta administrativa: usar sólo después de validar asistencia de la unidad. */
function validarM02RecoveryUnit(unidad, responsable, correo) {
 unidad=String(unidad||'').trim();
 const state=getMissionState(unidad);
 if (state.activeMission!=='M02') throw new Error('M02 no está pendiente para esta unidad.');
 const regs=getM02RecoveryRegistrations_().filter(r=>norm_(r.Unidad)===norm_(unidad));
 if (!regs.length) throw new Error('La unidad no tiene convocatoria de recuperación registrada.');
 const attended=regs.filter(r=>bool_(r.Asistencia));
 if (!attended.length) throw new Error('Primero marca Asistencia = SI en M02_Recuperacion_Convocatoria para las personas validadas.');
 completeRecovery_(unidad,'M02',responsable,correo,attended.length+' asistencias validadas');
 return getMissionState(unidad);
}

function uniqueEmails_(value) {
 const arr = Array.isArray(value) ? value : String(value || '').split(/[,;\s]+/);
 const map = {};
 arr.forEach(v => {
 const x = String(v||'').trim().toLowerCase();
 if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x)) map[x] = true;
 });
 return Object.keys(map);
}

function completeRecovery_(unidad, missionId, responsable, correo, detalle) {
 const sh = getOrCreateSheet_(MC.RECOVERY, recoveryHeaders_());
 const existing = getRecoveryMap_(unidad);
 if (existing[missionId]) return;
 appendObject_(sh, recoveryHeaders_(), {
 'Timestamp':now_(),'Unidad':unidad,'Mision_ID':missionId,'Estado':'CUMPLIDA',
 'Puntos_otorgados':MC.MISSIONS[missionId].recovery,
 'Responsable':String(responsable||''),'Correo':String(correo||''),
 'Detalle':String(detalle||'')
 });
}

/* =========================
 M03 · LABORATORIO SMA
 ========================= */

function m03SampleHeaders_() {
 return [
 'Timestamp','Unidad','Responsable','Correo',
 'Asesor_No','Asesor','Prospecto_No','Prospecto',
 'SMA',
 'Control_contacto','Situacion_objecion','Actividad_concreta',
 'Fecha_comprometida','Seguimiento_CRM','Feedback_avance',
 'Criterios_cumplidos','Calidad_pct'
 ];
}

function m03CloseHeaders_() {
 return [
 'Timestamp','Unidad','Responsable','Correo','Aprendizaje',
 'Total_SMA','Cumplimiento_criterios_pct','Principal_oportunidad',
 'Puntos_otorgados','Tipo'
 ];
}

function getM03Data_(unidad) {
 const key = norm_(unidad);
 const sampleSh = sheet_(MC.M03_SAMPLE);
 const closeSh = sheet_(MC.M03_CLOSE);
 const samples = sampleSh ? rowsAsObjects_(sampleSh).filter(r => norm_(r.Unidad) === key) : [];
 const closes = closeSh ? rowsAsObjects_(closeSh).filter(r => norm_(r.Unidad) === key) : [];

 const counts = {control:0,situacion:0,concreta:0,fecha:0,crm:0,feedback:0};
 samples.forEach(r => {
 if (bool_(r.Control_contacto)) counts.control++;
 if (bool_(r.Situacion_objecion)) counts.situacion++;
 if (bool_(r.Actividad_concreta)) counts.concreta++;
 if (bool_(r.Fecha_comprometida)) counts.fecha++;
 if (bool_(r.Seguimiento_CRM)) counts.crm++;
 if (bool_(r.Feedback_avance)) counts.feedback++;
 });

 const denom = samples.length || 1;
 const criteria = MC.SMA_CRITERIA.map(c => ({
 key:c.key, label:c.label, count:counts[c.key],
 pct:Math.round((counts[c.key]/denom)*100)
 }));
 criteria.sort((a,b)=>a.pct-b.pct);

 const totalChecks = Object.keys(counts).reduce((s,k)=>s+counts[k],0);
 const quality = samples.length ? Math.round(totalChecks/(samples.length*6)*100) : 0;

 return {
 completed: closes.length > 0,
 samples:samples.length,
 quality:quality,
 criteria:criteria,
 mainOpportunity:samples.length ? criteria[0].label : '',
 close:closes.length ? closes[closes.length-1] : null
 };
}

function submitM03(payload) {
 payload = payload || {};
 const unidad = String(payload.unidad || '').trim();
 const state = getMissionState(unidad);
 if (state.activeMission !== 'M03') throw new Error('M03 no está habilitada para esta unidad.');

 const advisors = Array.isArray(payload.advisors) ? payload.advisors : [];
 if (advisors.length !== 3) throw new Error('El laboratorio requiere exactamente 3 asesores.');

 const flat = [];
 const advisorNames = {};
 advisors.forEach((a, ai) => {
 const advisor = String(a.name || '').trim();
 if (!advisor) throw new Error('Captura el nombre del asesor ' + (ai+1) + '.');
 if (advisorNames[norm_(advisor)]) throw new Error('Los tres asesores deben ser diferentes.');
 advisorNames[norm_(advisor)] = true;

 const prospects = Array.isArray(a.prospects) ? a.prospects : [];
 if (prospects.length !== 2) throw new Error('Cada asesor requiere 2 prospectos.');

 prospects.forEach((p,pi) => {
 const prospect = String(p.prospect || '').trim();
 const sma = String(p.sma || '').trim();
 if (!prospect) throw new Error('Captura el prospecto ' + (pi+1) + ' del asesor ' + (ai+1) + '.');
 if (!sma) throw new Error('Captura la SMA del prospecto ' + (pi+1) + ' del asesor ' + (ai+1) + '.');

 const criteria = p.criteria || {};
 flat.push({
 advisorNo:ai+1, advisor:advisor, prospectNo:pi+1, prospect:prospect, sma:sma,
 criteria:{
 control:!!criteria.control, situacion:!!criteria.situacion,
 concreta:!!criteria.concreta, fecha:!!criteria.fecha,
 crm:!!criteria.crm, feedback:!!criteria.feedback
 }
 });
 });
 });

 if (flat.length !== 6) throw new Error('El laboratorio debe contener 6 SMA.');
 const learning = String(payload.learning || '').trim();
 if (learning.length < 10) throw new Error('Agrega una reflexión breve sobre el aprendizaje del muestreo.');

 const lock = LockService.getScriptLock();
 lock.waitLock(20000);
 try {
 const recheck = getMissionState(unidad);
 if (recheck.missions.M03.completed) throw new Error('M03 ya fue completada por esta unidad.');

 const sampleHeaders = m03SampleHeaders_();
 const sampleSh = getOrCreateSheet_(MC.M03_SAMPLE, sampleHeaders);

 const criteriaTotals = {control:0,situacion:0,concreta:0,fecha:0,crm:0,feedback:0};
 flat.forEach(x => {
 let n = 0;
 Object.keys(criteriaTotals).forEach(k => {
 if (x.criteria[k]) { criteriaTotals[k]++; n++; }
 });
 appendObject_(sampleSh, sampleHeaders, {
 'Timestamp':now_(),'Unidad':unidad,
 'Responsable':String(payload.responsable||''),'Correo':String(payload.correo||''),
 'Asesor_No':x.advisorNo,'Asesor':x.advisor,
 'Prospecto_No':x.prospectNo,'Prospecto':x.prospect,'SMA':x.sma,
 'Control_contacto':x.criteria.control?'SI':'NO',
 'Situacion_objecion':x.criteria.situacion?'SI':'NO',
 'Actividad_concreta':x.criteria.concreta?'SI':'NO',
 'Fecha_comprometida':x.criteria.fecha?'SI':'NO',
 'Seguimiento_CRM':x.criteria.crm?'SI':'NO',
 'Feedback_avance':x.criteria.feedback?'SI':'NO',
 'Criterios_cumplidos':n,'Calidad_pct':Math.round(n/6*100)
 });
 });

 const totalChecks = Object.keys(criteriaTotals).reduce((s,k)=>s+criteriaTotals[k],0);
 const quality = Math.round(totalChecks/(6*6)*100);
 const sorted = MC.SMA_CRITERIA.map(c => ({
 label:c.label, count:criteriaTotals[c.key]
 })).sort((a,b)=>a.count-b.count);

 const closeHeaders = m03CloseHeaders_();
 const closeSh = getOrCreateSheet_(MC.M03_CLOSE, closeHeaders);
 appendObject_(closeSh, closeHeaders, {
 'Timestamp':now_(),'Unidad':unidad,
 'Responsable':String(payload.responsable||''),'Correo':String(payload.correo||''),
 'Aprendizaje':learning,'Total_SMA':6,
 'Cumplimiento_criterios_pct':quality,
 'Principal_oportunidad':sorted[0].label,
 'Puntos_otorgados':MC.MISSIONS.M03.base,
 'Tipo':isExpired_('M03')?'RECUPERACION':'ACTUAL'
 });

 return getMissionState(unidad);
 } finally {
 lock.releaseLock();
 }
}

/* =========================
 MOTOR DE PROGRESIÓN
 ========================= */

function getMissionCatalog_() {
 const sh = sheet_(MC.MISSION_CATALOG);
 if (!sh) return {};

 const map = {};
 rowsAsObjects_(sh).forEach(r => {
   const id = norm_(r.Mision_ID);
   if (!id) return;

   map[id] = {
     id:id,
     name:String(r.Nombre || '').trim(),
     detail:String(r.Detalle_mision || '').trim(),
     recoveryInstructions:String(r.Instrucciones_recuperacion || '').trim(),
     recoveryPoints:Number(r.Puntos_recuperacion || 0)
   };
 });
 return map;
}

function getMissionState(unidad) {
 unidad = String(unidad || '').trim();
 if (!unidad) throw new Error('Selecciona una unidad.');

 const hist = getHistorical_(unidad);
 const rec = getRecoveryMap_(unidad);
 const m03 = getM03Data_(unidad);
 const m02Recovery = getM02RecoveryData_(unidad);
 const missionCatalog = getMissionCatalog_();

 const m01Done = hist.M01.completed || !!rec.M01;
 const m02Done = hist.M02.completed || !!rec.M02;
 const m03Done = m03.completed;

 const missions = {
 M01: missionView_('M01', hist.M01, rec.M01, m01Done),
 M02: missionView_('M02', hist.M02, rec.M02, m02Done),
 M03: {
 id:'M03', name:MC.MISSIONS.M03.name, deadline:MC.MISSIONS.M03.deadline,
 completed:m03Done, historical:false,
 points:m03Done ? Number(m03.close.Puntos_otorgados || 500) : 0,
 status:m03Done?'COMPLETED':'LOCKED',
 mode:m03Done ? String(m03.close.Tipo || 'ACTUAL') : null
 }
 };

 let activeMission = null;
 let activeMode = null;

 // Se respetan misiones históricas cumplidas aunque estén fuera de secuencia.
 // Sólo se exige regularizar la primera pendiente.
 if (!m01Done) {
 activeMission = 'M01'; activeMode = 'RECOVERY';
 } else if (!m02Done) {
 activeMission = 'M02'; activeMode = 'RECOVERY';
 } else if (!m03Done) {
 activeMission = 'M03'; activeMode = isExpired_('M03') ? 'RECOVERY' : 'CURRENT';
 }

 Object.keys(missions).forEach(id => {
 if (missions[id].completed) missions[id].status = 'COMPLETED';
 else if (id === activeMission) missions[id].status = activeMode === 'RECOVERY' ? 'RECOVERY' : 'CURRENT';
 else missions[id].status = 'LOCKED';
 });

 const historicalPoints = hist.points;
 const recoveryPoints = Object.keys(rec).reduce((s,k)=>s+Number(rec[k].points||0),0);
 const m03Points = m03Done ? Number(m03.close.Puntos_otorgados || 500) : 0;

 const activeMissionDetail = activeMission && missionCatalog[activeMission]
   ? missionCatalog[activeMission]
   : null;

 return {
 unidad:unidad,
 activeMission:activeMission,
 activeMode:activeMode,
 activeMissionDetail:activeMissionDetail,
 allCompleted:!activeMission,
 points:{
 historical:historicalPoints,
 recovery:recoveryPoints,
 m03:m03Points,
 total:historicalPoints+recoveryPoints+m03Points
 },
 missions:missions,
 m02Recovery:m02Recovery,
 m03:{
 samples:m03.samples,
 quality:m03.quality,
 criteria:m03.criteria,
 mainOpportunity:m03.mainOpportunity
 },
 smaCriteria:MC.SMA_CRITERIA,
 generatedAt:Utilities.formatDate(now_(),MC.TZ,'dd/MM/yyyy HH:mm')
 };
}

function missionView_(id, historical, recovery, done) {
 const cfg = MC.MISSIONS[id];
 return {
 id:id, name:cfg.name, deadline:cfg.deadline, completed:done,
 historical:!!(historical && historical.completed),
 recovered:!!recovery,
 points: historical && historical.completed ? Number(historical.points||0) :
 recovery ? Number(recovery.points||0) : 0,
 bonus: historical && historical.completed ? Number(historical.bonus||0) : 0,
 status:done?'COMPLETED':'LOCKED',
 mode: historical && historical.completed ? 'HISTORICAL' : recovery ? 'RECOVERY' : null
 };
}

/* =========================
 COMPATIBILIDAD BÁSICA
 ========================= */

// Mantiene un endpoint simple para código previo que espere getEstadoGeneral.
function getEstadoGeneral(unidad) {
 return getMissionState(unidad);
}

// Si tu proyecto actual ya tiene un directorio externo de comités,
// puedes reemplazar este fallback por tu función existente.
function getDirectorioComites_() {
 return [];
}
