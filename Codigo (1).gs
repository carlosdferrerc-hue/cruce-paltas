/**
 * Recibe las muestras de la app "Cruce Calidad Paltas" y las agrega a la planilla.
 * Proyecto independiente: crear en script.google.com > Nuevo proyecto.
 * No toca ningún otro script de la planilla.
 */
const ID_PLANILLA = "1qXii2FKVpsl-MVuu0wBgT-tYewVXoLhTbf9vePr490Y";  // cruce calidad paltas ACQ
const CLAVE = "cambia-esta-clave";             // la misma que pondrás en la app
const HOJA  = "Respuestas de formulario 1";    // hoja donde está la tabla Form_Responses
const HOJA_IDS = "_sync_ids";                  // hoja oculta para no duplicar envíos

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const datos = JSON.parse(e.postData.contents);
    if (datos.clave !== CLAVE) return json({ ok: false, error: "Clave incorrecta" });

    const ss = SpreadsheetApp.openById(ID_PLANILLA);
    const hoja = ss.getSheetByName(HOJA);
    if (!hoja) return json({ ok: false, error: "No existe la hoja " + HOJA });

    let hIds = ss.getSheetByName(HOJA_IDS);
    if (!hIds) { hIds = ss.insertSheet(HOJA_IDS); hIds.hideSheet(); }
    const yaEnviados = new Set(
      hIds.getLastRow() ? hIds.getRange(1, 1, hIds.getLastRow(), 1).getValues().flat() : []
    );

    const tz = ss.getSpreadsheetTimeZone();
    const filas = [], guardados = [], ids = [];
    (datos.registros || []).forEach(r => {
      if (yaEnviados.has(r.id)) { guardados.push(r.id); return; }   // ya estaba: se confirma sin duplicar
      filas.push([
        new Date(r.ts),          // Marca temporal (hora en que se tomó la muestra en el teléfono)
        r.nombre,                // Nombre
        r.linea,                 // LÍNEA SALIDA
        r.tipo,                  // Tipo Muestra
        r.categoria,             // Categoría
        valorCalibre(r.calibre), // CALIBRE (número si es número; si no, el texto tal cual; puede ir vacío)
        Number(r.count),         // COUNT
        Number(r.cat1),          // CANTIDAD CAT 1
        Number(r.cat1n),         // CANTIDAD CAT 1 NEGRA
        Number(r.cat2),          // CANTIDAD CAT 2
        Number(r.descarte),      // CANTIDAD DESCARTE
        r.comentario || ""       // COMENTARIO
      ]);
      guardados.push(r.id);
      ids.push([r.id, Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm:ss")]);
    });

    if (filas.length) {
      const fila = ultimaFilaConDatos(hoja) + 1;
      hoja.getRange(fila, 1, filas.length, 12).setValues(filas);
      hoja.getRange(fila, 1, filas.length, 1).setNumberFormat("dd/MM/yyyy H:mm:ss");
      hIds.getRange(hIds.getLastRow() + 1, 1, ids.length, 2).setValues(ids);
    }
    return json({ ok: true, guardados: guardados, nuevas: filas.length });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Para probar desde el navegador que la URL responde. */
function doGet() {
  return json({ ok: true, app: "Cruce Calidad Paltas" });
}

/** Última fila con dato en la columna A (ignora filas vacías con formato de la tabla). */
function ultimaFilaConDatos(hoja) {
  const col = hoja.getRange(1, 1, hoja.getMaxRows(), 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) if (col[i][0] !== "") return i + 1;
  return 1;
}

function valorCalibre(c) {
  c = String(c == null ? "" : c).trim();
  return /^\d+([.,]\d+)?$/.test(c) ? Number(c.replace(",", ".")) : c;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
