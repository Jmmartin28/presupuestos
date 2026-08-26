/* shg-cloud.js — Autenticación (MSAL.js) y acceso a las listas de SharePoint vía
   Microsoft Graph. Mismo patrón que la app de Daily Cost, con dos mejoras: se
   comprueba el código HTTP y se reintenta ante 429 (throttling), y se resuelven los
   nombres internos de columna (así "Año" u otros nombres con acentos funcionan). */
(function () {
  "use strict";
  var C = window.SHG_CONFIG;

  // En localhost la redirección de login vuelve al propio origen (hay que registrar
  // http://localhost:8000 como URI SPA en Azure); en producción, la de config.js.
  var _esLocalhost = (location.hostname === "localhost" || location.hostname === "127.0.0.1");
  var _redirectUri = _esLocalhost ? location.origin : (C.redirectUri || (location.origin + location.pathname));

  var msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: C.clientId,
      authority: "https://login.microsoftonline.com/" + C.tenant,
      redirectUri: _redirectUri,
    },
    cache: { cacheLocation: "localStorage" },
  });

  // ¿Debe activarse la nube/puerta? En producción (github.io) siempre; en local solo
  // en el puerto 8000 (pruebas con login real); otros puertos/ficheros = modo local.
  function _modoNube() {
    if (location.hostname.indexOf("github.io") !== -1) return true;
    return _esLocalhost && location.port === "8000";
  }

  var _account = null, _siteId = null, _listId = {}, _cols = {};

  // Procesa el retorno del login y recupera la cuenta si ya hay sesión.
  async function init() {
    var resp = await msalInstance.handleRedirectPromise();
    if (resp && resp.account) _account = resp.account;
    if (!_account) {
      var accs = msalInstance.getAllAccounts();
      if (accs.length) _account = accs[0];
    }
    return _account;
  }
  function login() { return msalInstance.loginRedirect({ scopes: C.scopes }); }
  function logout() { return msalInstance.logoutRedirect(); }
  function cuenta() { return _account; }

  async function getToken() {
    if (!_account) {
      var accs = msalInstance.getAllAccounts();
      if (accs.length) _account = accs[0];
    }
    try {
      var r = await msalInstance.acquireTokenSilent({ scopes: C.scopes, account: _account });
      return r.accessToken;
    } catch (e) {
      await msalInstance.acquireTokenRedirect({ scopes: C.scopes, account: _account });
      throw e;   // el redirect recarga la página
    }
  }

  // Llamada a Graph con control de estado y reintentos ante 429 (Retry-After).
  async function graphCall(url, method, body, token) {
    token = token || (await getToken());
    for (var intento = 0; intento < 4; intento++) {
      var opts = { method: method || "GET",
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" } };
      if (body) opts.body = JSON.stringify(body);
      var r = await fetch("https://graph.microsoft.com/v1.0" + url, opts);
      if (r.status === 429) {
        var wait = parseInt(r.headers.get("Retry-After") || "2", 10) || 2;
        await new Promise(function (s) { setTimeout(s, wait * 1000); });
        continue;
      }
      // 409 resourceModified (eTag): otro guardado toco el item entre lectura y escritura.
      // El PATCH de /fields es una fusion, asi que reintentar tras una pausa corta lo resuelve.
      if (r.status === 409 && method !== "POST" && intento < 3) {
        await new Promise(function (s) { setTimeout(s, 400 + intento * 500); });
        continue;
      }
      if (!r.ok) { var t = await r.text(); throw new Error("Graph " + r.status + " " + url + " :: " + t); }
      if (method === "PATCH" || method === "DELETE" || r.status === 204) return null;
      return r.json();
    }
    throw new Error("Graph 429 repetido en " + url);
  }

  async function getSiteId() {
    if (_siteId) return _siteId;
    var s = await graphCall("/sites/" + C.site);
    _siteId = s.id;
    return _siteId;
  }
  async function getListId(nombre) {
    if (_listId[nombre]) return _listId[nombre];
    var sid = await getSiteId();
    var r = await graphCall("/sites/" + sid + "/lists?$filter=displayName eq '" + nombre + "'&$select=id,displayName");
    if (!r.value || !r.value.length) throw new Error("No existe la lista: " + nombre);
    _listId[nombre] = r.value[0].id;
    return _listId[nombre];
  }
  // Mapa nombre visible -> nombre interno de columna (para escribir fields).
  async function getCols(nombre) {
    if (_cols[nombre]) return _cols[nombre];
    var sid = await getSiteId(), lid = await getListId(nombre);
    var r = await graphCall("/sites/" + sid + "/lists/" + lid + "/columns?$select=name,displayName");
    var map = {};
    (r.value || []).forEach(function (c) { map[c.displayName] = c.name; map[c.name] = c.name; });
    _cols[nombre] = map;
    return map;
  }
  function interno(map, nombre) { return map[nombre] || nombre; }

  async function leerItems(nombre) {
    var sid = await getSiteId(), lid = await getListId(nombre);
    var url = "/sites/" + sid + "/lists/" + lid + "/items?$expand=fields&$top=200";
    var out = [], token = await getToken();
    while (url) {
      var r = await graphCall(url, "GET", null, token);
      out = out.concat(r.value || []);
      url = r["@odata.nextLink"] ? r["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
    }
    return out;
  }
  async function crearItem(nombre, fields) {
    var sid = await getSiteId(), lid = await getListId(nombre);
    return graphCall("/sites/" + sid + "/lists/" + lid + "/items", "POST", { fields: fields });
  }
  async function actualizarItem(nombre, itemId, fields) {
    var sid = await getSiteId(), lid = await getListId(nombre);
    return graphCall("/sites/" + sid + "/lists/" + lid + "/items/" + itemId + "/fields", "PATCH", fields);
  }
  async function borrarItem(nombre, itemId) {
    var sid = await getSiteId(), lid = await getListId(nombre);
    return graphCall("/sites/" + sid + "/lists/" + lid + "/items/" + itemId, "DELETE");
  }

  // ---- Capa de datos: versiones (con sus datos por hotel) sobre las dos listas ----
  var LS_VS = "shg_versiones", LS_ACT = "shg_version_activa";
  var _conectado = false;
  var _idByTitle = {};   // { nombreLista: { Title: itemId } } para saber POST vs PATCH
  var _pending = {};     // guardados en espera hasta tener conexión (clave -> [tipo, ...args])

  function _parse(s) { try { return s ? JSON.parse(s) : {}; } catch (e) { return {}; } }

  // Al establecer la conexión, ejecuta los guardados que quedaron pendientes.
  function _flush() {
    _conectado = true;
    var p = _pending; _pending = {};
    Object.keys(p).forEach(function (k) {
      var x = p[k];
      if (x[0] === "version") pushVersion(x[1]); else pushHotel(x[1], x[2]);
    });
  }

  // Pinta el estado de la nube en el indicador visible (#abNube), si existe.
  function _estado(txt, tipo) {
    var e = document.getElementById("abNube");
    if (e) { e.textContent = "Nube: " + txt; e.className = "ab-nube " + (tipo || ""); }
  }

  // Descarga las dos listas y reconstruye el objeto shg_versiones en localStorage.
  async function cargarTodo() {
    var colsV = await getCols(C.listaVersiones), fAnio = interno(colsV, "Año");
    var vers = await leerItems(C.listaVersiones);
    var hots = await leerItems(C.listaHoteles);
    _idByTitle[C.listaVersiones] = {}; _idByTitle[C.listaHoteles] = {};
    var out = {};
    vers.forEach(function (it) {
      var f = it.fields || {}, id = f.Title; if (!id) return;
      _idByTitle[C.listaVersiones][id] = it.id;
      var hlist = null; try { var hj = JSON.parse(f.HotelesJSON || "null"); if (Array.isArray(hj)) hlist = hj; } catch (e) {}
      var mAud = (f.MesAuditado === "" || f.MesAuditado == null || isNaN(Number(f.MesAuditado))) ? null : Number(f.MesAuditado);
      var cMf = null; try { var cj = JSON.parse(f.CentralMfJSON || "null"); if (Array.isArray(cj) && cj.length === 12) cMf = cj; } catch (e) {}
      out[id] = { anio: +(f[fAnio] || 0), nombre: f.Nombre || "", autor: f.Autor || "",
        hipotesis: f.Hipotesis || "", incrementos: _parse(f.IncrementosJSON), hoteles: hlist, mesAuditado: mAud, centralMf: cMf,
        overrides: {}, medidas: {}, alojamiento: {}, personal: {}, pptoCat: {},
        creada: f.Creada || "", modificada: f.Modificada || "" };
    });
    hots.forEach(function (it) {
      var f = it.fields || {}, t = f.Title; if (!t) return;
      _idByTitle[C.listaHoteles][t] = it.id;
      var vid = f.VersionId, hid = String(f.HotelId);
      if (!out[vid]) return;
      if (f.OverridesJSON)    out[vid].overrides[hid]   = _parse(f.OverridesJSON);
      if (f.MedidasJSON)      out[vid].medidas[hid]      = _parse(f.MedidasJSON);
      if (f.AlojamientoJSON)  out[vid].alojamiento[hid]  = _parse(f.AlojamientoJSON);
      if (f.PersonalJSON)     out[vid].personal[hid]     = _parse(f.PersonalJSON);
      if (f.PptoCatJSON)      out[vid].pptoCat[hid]      = _parse(f.PptoCatJSON);
    });
    localStorage.setItem(LS_VS, JSON.stringify(out));
    _flush();
    return out;
  }

  // Lee solo los Title→id de las dos listas (para poder guardar), SIN tocar el
  // localStorage. Se usa al navegar entre pantallas para no pisar lo calculado.
  async function cargarIndice() {
    var sid = await getSiteId(), tok = await getToken();
    for (var i = 0; i < 2; i++) {
      var lista = i === 0 ? C.listaVersiones : C.listaHoteles;
      var lid = await getListId(lista);
      _idByTitle[lista] = {};
      var url = "/sites/" + sid + "/lists/" + lid + "/items?$select=id&$expand=fields($select=Title)&$top=500";
      while (url) {
        var r = await graphCall(url, "GET", null, tok);
        (r.value || []).forEach(function (it) { var t = it.fields && it.fields.Title; if (t) _idByTitle[lista][t] = it.id; });
        url = r["@odata.nextLink"] ? r["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
      }
    }
    _flush();
  }

  // Guarda (upsert) la fila de metadatos de una versión desde localStorage.
  async function pushVersion(versionId) {
    if (!_conectado) { _pending["v:" + versionId] = ["version", versionId]; _estado("guardando (esperando conexión)…", ""); return; }
    _estado("guardando…", "");
    try {
      var reg = (_parse(localStorage.getItem(LS_VS)))[versionId]; if (!reg) return;
      var cols = await getCols(C.listaVersiones), fAnio = interno(cols, "Año");
      var fields = { Title: versionId, Nombre: reg.nombre || "", Autor: reg.autor || "",
        Hipotesis: reg.hipotesis || "", IncrementosJSON: JSON.stringify(reg.incrementos || {}),
        Creada: reg.creada || new Date().toISOString(), Modificada: reg.modificada || new Date().toISOString() };
      fields[fAnio] = Number(reg.anio) || 0;
      if ("HotelesJSON" in cols) fields.HotelesJSON = JSON.stringify(reg.hoteles || []);   // solo si existe la columna
      if ("MesAuditado" in cols) fields.MesAuditado = (reg.mesAuditado == null ? "" : String(reg.mesAuditado));   // "" = automático
      if ("CentralMfJSON" in cols) fields.CentralMfJSON = JSON.stringify(reg.centralMf || null);   // override MF Central
      var id = _idByTitle[C.listaVersiones] && _idByTitle[C.listaVersiones][versionId];
      if (id) await actualizarItem(C.listaVersiones, id, fields);
      else { var c = await crearItem(C.listaVersiones, fields); _idByTitle[C.listaVersiones][versionId] = c.id; }
      _estado("guardado ✓ " + new Date().toLocaleTimeString(), "ok");
    } catch (e) { console.error("Error guardando la versión en SharePoint:", e); _estado("⚠ NO guardado: " + e.message, "err"); }
  }

  // Guarda (upsert) la fila de un hotel de una versión desde localStorage.
  async function pushHotel(versionId, hotelId) {
    if (!_conectado) { _pending["h:" + versionId + ":" + hotelId] = ["hotel", versionId, hotelId]; _estado("guardando (esperando conexión)…", ""); return; }
    _estado("guardando…", "");
    try {
      var reg = (_parse(localStorage.getItem(LS_VS)))[versionId]; if (!reg) return;
      var hid = String(hotelId), title = versionId + "_" + hid;
      var fields = { Title: title, VersionId: versionId, HotelId: Number(hotelId) || 0,
        OverridesJSON:   JSON.stringify((reg.overrides   && reg.overrides[hid])   || {}),
        MedidasJSON:     JSON.stringify((reg.medidas     && reg.medidas[hid])     || {}),
        AlojamientoJSON: JSON.stringify((reg.alojamiento && reg.alojamiento[hid]) || {}),
        PersonalJSON:    JSON.stringify((reg.personal    && reg.personal[hid])    || {}),
        PptoCatJSON:     JSON.stringify((reg.pptoCat     && reg.pptoCat[hid])     || {}) };
      var id = _idByTitle[C.listaHoteles] && _idByTitle[C.listaHoteles][title];
      if (id) await actualizarItem(C.listaHoteles, id, fields);
      else { var c = await crearItem(C.listaHoteles, fields); _idByTitle[C.listaHoteles][title] = c.id; }
      _estado("guardado ✓ " + new Date().toLocaleTimeString(), "ok");
    } catch (e) { console.error("Error guardando el hotel en SharePoint:", e); _estado("⚠ NO guardado: " + e.message, "err"); }
  }

  // Borra una versión por completo en SharePoint: su fila de metadatos y TODAS sus
  // filas por hotel. En modo local (sin SharePoint) no hace nada; el borrado local lo
  // realiza la pantalla de Inicio. En nube requiere conexión (si no, no se puede borrar
  // en remoto y volvería a aparecer en la siguiente sincronización).
  async function deleteVersion(versionId) {
    if (!_modoNube()) return;
    if (!_conectado) throw new Error("Sin conexión con SharePoint; reinténtalo cuando termine de cargar.");
    if (!_idByTitle[C.listaVersiones] || !_idByTitle[C.listaHoteles]) await cargarIndice();
    var vid = _idByTitle[C.listaVersiones] && _idByTitle[C.listaVersiones][versionId];
    if (vid) { await borrarItem(C.listaVersiones, vid); delete _idByTitle[C.listaVersiones][versionId]; }
    var mapH = _idByTitle[C.listaHoteles] || {}, prefijo = versionId + "_";
    var titulos = Object.keys(mapH).filter(function (t) {
      return t.indexOf(prefijo) === 0 && /^\d+$/.test(t.slice(prefijo.length));   // versionId_<idHotel>
    });
    for (var i = 0; i < titulos.length; i++) {
      await borrarItem(C.listaHoteles, mapH[titulos[i]]);
      delete mapH[titulos[i]];
    }
    _estado("versión borrada ✓ " + new Date().toLocaleTimeString(), "ok");
  }

  // Sincroniza una vez por sesión al abrir cualquier pantalla en GitHub Pages: hace
  // login, descarga las versiones a localStorage y recarga (para que la página se
  // pinte con los datos frescos). En local no hace nada (modo localStorage).
  async function autoSync(pagina) {
    if (!_modoNube()) {                                          // modo local (sin login ni puerta)
      setTimeout(function () { _estado("modo local (sin SharePoint)", ""); }, 0);
      return;
    }
    // Puerta de acceso: login corporativo + lista blanca + permiso de página.
    document.documentElement.style.visibility = "hidden";
    var ok;
    try { ok = await puerta(pagina); }
    catch (e) { console.error("Puerta:", e); _bloquear("No se pudo comprobar el acceso: " + e.message); return; }
    if (!ok) return;                                             // bloqueado, redirigido o yendo a login
    document.documentElement.style.visibility = "";
    _avisarSesion();                                            // las páginas re-filtran por permisos

    // Sincroniza las versiones (metadatos y datos por hotel) a localStorage. Ya no
    // hace falta recargar: la página se pinta cuando SHG.datos() entrega los datos.
    var primera = !sessionStorage.getItem("shg_synced");
    try {
      if (primera) { await cargarTodo(); sessionStorage.setItem("shg_synced", "1"); }
      else { await cargarIndice(); }
    } catch (e) { console.error("Nube (versiones):", e); }

    // Descarga los datos base de la página desde SharePoint y los entrega a la página.
    _loading(true);
    try {
      var D = await descargarDatosSP(pagina);
      _entregarDatos(D);
      setTimeout(function () { _loading(false); }, 60);   // se quita tras pintar
      _estado("conectado ✓ · " + (sesion() ? sesion().nombre : ""), "ok");
    } catch (e) {
      console.error("Nube (datos):", e);
      _loading(false);
      _estado("error: " + e.message, "err");
      _bloquear("No se pudieron cargar los datos: " + e.message);
    }
  }
  function conectado() { return _conectado; }

  // ---- Datos base de cada página (en la nube, o en local desde ./datos) ----
  var _datosResolve = null;
  var _datosPromesa = new Promise(function (res) { _datosResolve = res; });
  function _entregarDatos(D) { if (_datosResolve) { _datosResolve(D); _datosResolve = null; } }

  // Descarga datos/<pagina>.json de la biblioteca Presupuesto_Datos (SharePoint).
  // Se pide siempre fresco (sin caché): al republicar datos, la app los ve al instante.
  async function descargarDatosSP(pagina) {
    var sid = await getSiteId();
    var dr = await graphCall("/sites/" + sid + "/drives?$select=id,name");
    var drive = (dr.value || []).filter(function (d) { return d.name === C.listaDatos; })[0];
    if (!drive) throw new Error("No existe la biblioteca de datos: " + C.listaDatos);
    var tok = await getToken();
    // cache-buster para saltar cachés intermedias del navegador/proxy.
    var r = await fetch("https://graph.microsoft.com/v1.0/drives/" + drive.id +
      "/root:/" + pagina + ".json:/content?t=" + Date.now(),
      { headers: { Authorization: "Bearer " + tok }, cache: "no-store" });
    if (!r.ok) throw new Error("No se pudo descargar " + pagina + ".json (" + r.status + ")");
    return r.json();
  }

  // La usa cada página: devuelve los datos base. En local, del fichero local; en la
  // nube, los entrega autoSync tras la puerta y la descarga.
  function datos(pagina) {
    if (!_modoNube()) {
      return fetch("datos/" + pagina + ".json").then(function (r) {
        if (!r.ok) throw new Error("No se encontró datos/" + pagina + ".json"); return r.json();
      });
    }
    return _datosPromesa;
  }

  // ---- Permisos de usuario (lista blanca Presupuesto_Usuarios) ----
  var _sesion = null;              // permisos resueltos del usuario logueado
  var PAGINAS_TODAS = "todas";

  function _emailCuenta() {
    var a = _account || (msalInstance.getAllAccounts()[0]);
    return ((a && (a.username || a.name)) || "").toLowerCase();
  }
  function _partes(v) {
    return String(v == null ? "" : v).split(/[;,]/).map(function (s) { return s.trim(); }).filter(Boolean);
  }
  function _esVerdadero(v) {        // columna Sí/No: llega como true/false, "1"/"0", "Sí"/"No"
    if (v === true) return true;
    if (v === false || v == null) return false;
    var s = String(v).toLowerCase();
    return s === "1" || s === "true" || s === "sí" || s === "si" || s === "yes";
  }

  // Lee Presupuesto_Usuarios y resuelve los permisos del usuario actual (se cachea).
  async function cargarSesion() {
    if (_sesion) return _sesion;
    var email = _emailCuenta();
    var fila = null;
    try {
      var items = await leerItems(C.listaUsuarios);
      items.forEach(function (it) {
        var f = it.fields || {};
        if (String(f.Title || "").trim().toLowerCase() === email) fila = f;
      });
    } catch (e) { console.error("No se pudo leer la lista de usuarios:", e); }
    if (!fila) { _sesion = { email: email, autorizado: false }; return _sesion; }
    var pag = String(fila.Paginas || "todas").toLowerCase();
    _sesion = {
      email: email,
      nombre: fila.Nombre || email,
      autorizado: _esVerdadero(fila.Activo === undefined ? true : fila.Activo),
      ambito: String(fila.Ambito || "todos").toLowerCase(),   // todos | zona | hoteles
      zonas: _partes(fila.Zonas).map(function (z) { return z.toLowerCase(); }),
      hoteles: _partes(fila.Hoteles).map(Number).filter(function (n) { return !isNaN(n); }),
      paginas: pag.indexOf("todas") >= 0 ? PAGINAS_TODAS : _partes(pag),
    };
    return _sesion;
  }
  function sesion() { return _sesion; }
  // Avisa a las páginas de que ya se conocen los permisos, para que re-filtren/re-render.
  function _avisarSesion() { try { window.dispatchEvent(new Event("shg:sesion")); } catch (e) {} }

  function paginaPermitida(pagina) {
    if (!_sesion || !_sesion.autorizado) return false;
    return _sesion.paginas === PAGINAS_TODAS || _sesion.paginas.indexOf(pagina) >= 0;
  }
  // ¿El usuario puede ver este hotel? (idZona/nombreZona para el ámbito por zona)
  function hotelPermitido(idHotel, idZona, nombreZona) {
    var s = _sesion; if (!s || !s.autorizado) return false;
    if (s.ambito === "todos") return true;
    if (s.ambito === "hoteles") return s.hoteles.indexOf(Number(idHotel)) >= 0;
    if (s.ambito === "zona") {
      return s.zonas.indexOf(String(idZona).toLowerCase()) >= 0
          || s.zonas.indexOf(String(nombreZona || "").toLowerCase()) >= 0;
    }
    return false;
  }

  // Overlay de bloqueo a pantalla completa.
  // Indicador de carga a pantalla completa mientras se descargan los datos.
  function _loading(on) {
    var id = "shgCargando", e = document.getElementById(id);
    if (on) {
      if (!e) {
        e = document.createElement("div"); e.id = id;
        e.style.cssText = "position:fixed;inset:0;z-index:9998;display:flex;align-items:center;"
          + "justify-content:center;background:var(--bg,#f6f7f9);color:var(--muted,#6b7688);"
          + "font-family:'Segoe UI',system-ui,sans-serif;font-size:14px";
        e.textContent = "Cargando datos…";
        document.body.appendChild(e);
      }
    } else if (e) { e.remove(); }
  }

  function _bloquear(mensaje) {
    document.documentElement.style.visibility = "";
    var d = document.createElement("div");
    d.style.cssText = "position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;"
      + "background:#0f141b;color:#e7ecf3;font-family:'Segoe UI',system-ui,sans-serif;text-align:center;padding:24px";
    d.innerHTML = "<div><div style='font-size:40px;margin-bottom:12px'>🔒</div>"
      + "<div style='font-size:16px;max-width:440px;line-height:1.5'>" + mensaje + "</div>"
      + "<div style='margin-top:16px'><button id='shgSalir' style='font:inherit;padding:8px 16px;border-radius:8px;"
      + "border:1px solid #26303d;background:#161c26;color:#e7ecf3;cursor:pointer'>Cerrar sesión</button></div></div>";
    document.body.appendChild(d);
    var b = document.getElementById("shgSalir"); if (b) b.addEventListener("click", function () { logout(); });
  }

  // Puerta de acceso: login + lista blanca + permiso de página. Devuelve true si pasa.
  async function puerta(paginaActual) {
    await init();
    if (!cuenta()) { login(); return false; }
    var s = await cargarSesion();
    if (!s.autorizado) {
      _bloquear("Tu cuenta (" + s.email + ") no tiene acceso a esta aplicación. Contacta con el administrador.");
      return false;
    }
    if (paginaActual && !paginaPermitida(paginaActual)) {
      var destino = (s.paginas !== PAGINAS_TODAS && s.paginas[0]) || null;
      if (destino && destino !== paginaActual) { location.href = destino + ".html"; return false; }
      _bloquear("No tienes acceso a esta sección."); return false;
    }
    return true;
  }

  // ---- Exportación a Excel (.xlsx real, con formato, sin librerías externas) ----
  var _ENC = new TextEncoder();
  function _xmlesc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function _colName(i) {                 // 0 -> A, 25 -> Z, 26 -> AA
    var s = ""; i += 1;
    while (i > 0) { var r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = (i - 1 - r) / 26; }
    return s;
  }
  function _num(txt) {                    // número español -> Number, o null si es texto / % / €
    if (!txt || /[a-zA-Z]/.test(txt) || txt.indexOf("%") >= 0 || txt.indexOf("€") >= 0) return null;
    var t = txt.replace(/\s/g, "");
    if (t === "" || t === "-" || !/^-?[\d.]*(,\d+)?$/.test(t)) return null;
    var n = parseFloat(t.replace(/\./g, "").replace(",", "."));
    return isNaN(n) ? null : n;
  }
  var _CRC = (function () { var t = []; for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function _crc32(b) { var c = -1; for (var i = 0; i < b.length; i++) c = (c >>> 8) ^ _CRC[(c ^ b[i]) & 0xFF]; return (c ^ -1) >>> 0; }
  function _cat(arrs) { var n = 0, i; for (i = 0; i < arrs.length; i++) n += arrs[i].length; var o = new Uint8Array(n), p = 0; for (i = 0; i < arrs.length; i++) { o.set(arrs[i], p); p += arrs[i].length; } return o; }
  function _zip(files) {                  // files: [{name, data:Uint8Array}] -> Blob .xlsx (método store)
    var u16 = function (n) { return [n & 255, (n >> 8) & 255]; };
    var u32 = function (n) { return [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]; };
    var local = [], central = [], off = 0;
    files.forEach(function (f) {
      var nm = _ENC.encode(f.name), d = f.data, crc = _crc32(d);
      var lh = new Uint8Array([].concat([0x50, 0x4b, 0x03, 0x04], u16(20), u16(0), u16(0), u16(0), u16(0), u32(crc), u32(d.length), u32(d.length), u16(nm.length), u16(0)));
      local.push(lh, nm, d);
      central.push(new Uint8Array([].concat([0x50, 0x4b, 0x01, 0x02], u16(20), u16(20), u16(0), u16(0), u16(0), u16(0), u32(crc), u32(d.length), u32(d.length), u16(nm.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(off))), nm);
      off += lh.length + nm.length + d.length;
    });
    var cd = _cat(central);
    var end = new Uint8Array([].concat([0x50, 0x4b, 0x05, 0x06], u16(0), u16(0), u16(files.length), u16(files.length), u32(cd.length), u32(off), u16(0)));
    return new Blob([_cat(local), cd, end], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }
  // Color computado (rgb/rgba) -> ARGB "FFRRGGBB"; null si es transparente. Los tintes con
  // opacidad (rgba con alpha<1) se mezclan sobre blanco, para que salgan claros como en la app.
  function _argb(c) {
    var m = String(c).match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/); if (!m) return null;
    var a = m[4] === undefined ? 1 : parseFloat(m[4]); if (a === 0) return null;
    var h = function (n) { n = Math.round(a * (+n) + (1 - a) * 255); return (n < 16 ? "0" : "") + n.toString(16); };
    return ("FF" + h(m[1]) + h(m[2]) + h(m[3])).toUpperCase();
  }
  // Texto de celda -> número + formato (con sufijo € / % / pp) o texto.
  function _parseCell(txt) {
    txt = (txt || "").trim(); if (!txt) return { text: "" };
    var suf = /%/.test(txt) ? "%" : (/€/.test(txt) ? "€" : (/\d\s*pp\b/.test(txt) ? "pp" : ""));
    var t = txt.replace(/%|€|pp/g, "").replace(/\s/g, "");
    if (/[a-zA-Z]/.test(t) || t === "" || t === "-" || !/^-?[\d.]*(,\d+)?$/.test(t)) return { text: txt };
    var n = parseFloat(t.replace(/\./g, "").replace(",", ".")); if (isNaN(n)) return { text: txt };
    var fmt = (t.indexOf(",") >= 0 ? "#,##0.0" : "#,##0");
    if (suf === "%") fmt += "&quot;%&quot;"; else if (suf === "€") fmt += "&quot; €&quot;"; else if (suf === "pp") fmt += "&quot; pp&quot;";
    return { num: n, fmt: fmt };
  }
  // Estado de estilos de un libro (compartido por todas sus hojas).
  function _nuevoLibro() {
    var L = {
      fonts: ['<font><sz val="10"/><name val="Calibri"/></font>'], fontKey: { "|0": 0 },
      fills: ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], fillKey: {},
      numFmts: [], numKey: {}, nextNum: 164,
      xfs: ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'], xfKey: {},
      hojas: [], usadas: {}
    };
    L.fontId = function (color, bold) { var k = (color || "") + "|" + (bold ? 1 : 0); if (L.fontKey[k] != null) return L.fontKey[k];
      L.fonts.push("<font>" + (bold ? "<b/>" : "") + '<sz val="10"/>' + (color ? '<color rgb="' + color + '"/>' : "") + '<name val="Calibri"/></font>'); return L.fontKey[k] = L.fonts.length - 1; };
    L.fillId = function (bg) { if (!bg) return 0; if (L.fillKey[bg] != null) return L.fillKey[bg];
      L.fills.push('<fill><patternFill patternType="solid"><fgColor rgb="' + bg + '"/></patternFill></fill>'); return L.fillKey[bg] = L.fills.length - 1; };
    L.numId = function (fmt) { if (!fmt) return 0; if (L.numKey[fmt] != null) return L.numKey[fmt];
      var id = L.nextNum++; L.numFmts.push('<numFmt numFmtId="' + id + '" formatCode="' + fmt + '"/>'); return L.numKey[fmt] = id; };
    L.xfId = function (fi, fl, nm, align) { var k = fi + "|" + fl + "|" + nm + "|" + align; if (L.xfKey[k] != null) return L.xfKey[k];
      L.xfs.push('<xf numFmtId="' + nm + '" fontId="' + fi + '" fillId="' + fl + '" borderId="0" applyBorder="1"' + (nm ? ' applyNumberFormat="1"' : "") + (fi ? ' applyFont="1"' : "") + (fl ? ' applyFill="1"' : "") + '><alignment horizontal="' + align + '"/></xf>'); return L.xfKey[k] = L.xfs.length - 1; };
    return L;
  }
  function _stylesXml(L) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (L.numFmts.length ? '<numFmts count="' + L.numFmts.length + '">' + L.numFmts.join("") + "</numFmts>" : "") +
      '<fonts count="' + L.fonts.length + '">' + L.fonts.join("") + "</fonts>" +
      '<fills count="' + L.fills.length + '">' + L.fills.join("") + "</fills>" +
      '<borders count="1"><border><left/><right/><top/><bottom style="thin"><color rgb="FFDDDDDD"/></bottom><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + L.xfs.length + '">' + L.xfs.join("") + "</cellXfs>" +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  }
  // Construye el XML de UNA hoja leyendo el estilo real de cada celda (usa el estado de estilos L).
  function _hojaXml(L, el, opts) {
    var fontId = L.fontId, fillId = L.fillId, numId = L.numId, xfId = L.xfId;
    var filas = [].slice.call(el.querySelectorAll("tr")), xml = "", nf = 0, merges = [], mesCols = {}, maxCol = 0, pend = {};
    filas.forEach(function (tr) {
      var celdas = [].slice.call(tr.children).filter(function (c) { return c.tagName === "TD" || c.tagName === "TH"; });
      if (!celdas.length) return;
      var esCab = tr.parentNode && tr.parentNode.tagName === "THEAD";
      nf += 1; var ci = 0; xml += '<row r="' + nf + '">';
      celdas.forEach(function (cel) {
        while (pend[ci] > 0) ci++;                        // saltar columnas ocupadas por un rowspan
        var span = parseInt(cel.getAttribute("colspan") || "1", 10) || 1;
        var rspan = parseInt(cel.getAttribute("rowspan") || "1", 10) || 1;
        var ref = _colName(ci) + nf, cs = getComputedStyle(cel);
        var bg = _argb(cs.backgroundColor), col = _argb(cs.color);
        var bold = (+cs.fontWeight) >= 600 || cs.fontWeight === "bold";
        var inp = cel.querySelector("input"), txt;
        if (inp) { txt = inp.value; }
        else { var cc = cel.cloneNode(true); cc.querySelectorAll(".delta,.pctv,.cx").forEach(function (x) { x.remove(); }); txt = (cc.textContent || "").trim(); }
        var esLbl = cel.classList.contains("lbl") || (ci === 0 && !esCab);
        var p = (esCab || esLbl) ? { text: txt } : _parseCell(txt);
        var align = esCab ? "center" : (p.num != null ? "right" : (esLbl ? "left" : "right"));
        var s = xfId(fontId(col, bold || esCab), fillId(bg), (p.num != null ? numId(p.fmt) : 0), align);
        if (p.num != null) xml += '<c r="' + ref + '" s="' + s + '"><v>' + p.num + "</v></c>";
        else xml += '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t>' + _xmlesc(txt) + "</t></is></c>";
        if (span > 1 || rspan > 1) merges.push(ref + ":" + _colName(ci + span - 1) + (nf + rspan - 1));   // combinadas
        if (rspan > 1) for (var j = 0; j < span; j++) pend[ci + j] = rspan;
        if (cel.classList.contains("mes")) for (var k = 0; k < span; k++) mesCols[ci + k] = 1;
        ci += span;
      });
      if (ci > maxCol) maxCol = ci;
      Object.keys(pend).forEach(function (c) { if (pend[c] > 0) pend[c]--; });   // fila consumida
      xml += "</row>";
    });
    // Anchos de columna (concepto ancho; datos algo más anchos; meses agrupados si opts.group).
    var colsArr = ['<col min="1" max="1" width="26"/>'], outlineCol = 0;
    var run = null;
    var flush = function (end) { if (run === null) return; var grp = !!(opts && opts.group && run.mes);
      colsArr.push('<col min="' + (run.ini + 1) + '" max="' + (end + 1) + '" width="' + (grp ? 10.5 : 12) + '"' + (grp ? ' outlineLevel="1" hidden="1"' : "") + "/>");
      if (grp) outlineCol = 1; run = null; };
    for (var g = 1; g < maxCol; g++) {
      var isMes = !!mesCols[g];
      if (run === null) run = { mes: isMes, ini: g };
      else if (run.mes !== isMes) { flush(g - 1); run = { mes: isMes, ini: g }; }
    }
    flush(maxCol - 1);
    var colsXml = colsArr.join("");
    var mergesXml = merges.length ? '<mergeCells count="' + merges.length + '">' +
      merges.map(function (m) { return '<mergeCell ref="' + m + '"/>'; }).join("") + '</mergeCells>' : "";
    var fz = opts && opts.freeze;   // {x, y}: inmoviliza x columnas y y filas (paneles)
    var sheetViews = fz ? '<sheetViews><sheetView workbookViewId="0">' +
      '<pane xSplit="' + fz.x + '" ySplit="' + fz.y + '" topLeftCell="' + _colName(fz.x) + (fz.y + 1) + '" activePane="bottomRight" state="frozen"/>' +
      '</sheetView></sheetViews>' : "";
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (outlineCol ? '<sheetPr><outlinePr summaryRight="1"/></sheetPr>' : "") +
      sheetViews +
      (outlineCol ? '<sheetFormatPr defaultRowHeight="15" outlineLevelCol="1"/>' : "") +
      (colsXml ? "<cols>" + colsXml + "</cols>" : "") +
      "<sheetData>" + xml + "</sheetData>" + mergesXml + "</worksheet>";
  }
  // Añade una hoja al libro AHORA (lee los estilos vivos de la tabla en este momento).
  // Se llama repetidamente renderizando cada hotel antes de cada llamada.
  function agregarHoja(L, nombre, el, opts) {
    el = typeof el === "string" ? document.getElementById(el) : el; if (!el) return;
    var nm = String(nombre || "Hoja").replace(/[:\\\/?*\[\]]/g, " ").trim().slice(0, 31) || "Hoja";
    while (L.usadas[nm.toLowerCase()]) nm = nm.slice(0, 27) + " (" + (L.hojas.length + 1) + ")";
    L.usadas[nm.toLowerCase()] = 1;
    L.hojas.push({ nombre: nm, xml: _hojaXml(L, el, opts) });
  }
  // Ensambla y descarga el libro con las hojas ya añadidas.
  function descargarLibro(L, nombre) {
    var enc = function (s) { return _ENC.encode(s); };
    var partes = [
      { name: "_rels/.rels", data: enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>') },
      { name: "xl/styles.xml", data: enc(_stylesXml(L)) }
    ];
    var ctOv = "", wbSheets = "", wbRel = "";
    L.hojas.forEach(function (s, i) {
      var n = i + 1;
      partes.push({ name: "xl/worksheets/sheet" + n + ".xml", data: enc(s.xml) });
      ctOv += '<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      wbSheets += '<sheet name="' + _xmlesc(s.nombre) + '" sheetId="' + n + '" r:id="rId' + n + '"/>';
      wbRel += '<Relationship Id="rId' + n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/>';
    });
    partes.push({ name: "[Content_Types].xml", data: enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + ctOv + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>') });
    partes.push({ name: "xl/workbook.xml", data: enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + wbSheets + '</sheets></workbook>') });
    partes.push({ name: "xl/_rels/workbook.xml.rels", data: enc('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + wbRel + '<Relationship Id="rId' + (L.hojas.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>') });
    var url = URL.createObjectURL(_zip(partes)), a = document.createElement("a");
    a.href = url;
    a.download = (nombre || "informe") + "_" + new Date().toISOString().slice(0, 10).replace(/-/g, "") + ".xlsx";
    document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1500);
  }
  function exportarExcel(tabla, nombre, opts) {
    var L = _nuevoLibro(); agregarHoja(L, (opts && opts.hoja) || "Datos", tabla, opts); descargarLibro(L, nombre);
  }

  window.SHG = {
    init: init, login: login, logout: logout, cuenta: cuenta,
    exportarExcel: exportarExcel, nuevoLibro: _nuevoLibro, agregarHoja: agregarHoja, descargarLibro: descargarLibro,
    graphCall: graphCall, getSiteId: getSiteId, getListId: getListId,
    getCols: getCols, interno: interno,
    leerItems: leerItems, crearItem: crearItem, actualizarItem: actualizarItem, borrarItem: borrarItem,
    cargarTodo: cargarTodo, pushVersion: pushVersion, pushHotel: pushHotel, deleteVersion: deleteVersion,
    autoSync: autoSync, conectado: conectado,
    cargarSesion: cargarSesion, sesion: sesion, puerta: puerta,
    paginaPermitida: paginaPermitida, hotelPermitido: hotelPermitido,
    datos: datos,
  };
})();
