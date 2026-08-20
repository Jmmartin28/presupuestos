/* shg-cloud.js — Autenticación (MSAL.js) y acceso a las listas de SharePoint vía
   Microsoft Graph. Mismo patrón que la app de Daily Cost, con dos mejoras: se
   comprueba el código HTTP y se reintenta ante 429 (throttling), y se resuelven los
   nombres internos de columna (así "Año" u otros nombres con acentos funcionan). */
(function () {
  "use strict";
  var C = window.SHG_CONFIG;

  var msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: C.clientId,
      authority: "https://login.microsoftonline.com/" + C.tenant,
      redirectUri: C.redirectUri || (location.origin + location.pathname),
    },
    cache: { cacheLocation: "localStorage" },
  });

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

  function _parse(s) { try { return s ? JSON.parse(s) : {}; } catch (e) { return {}; } }

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
      out[id] = { anio: +(f[fAnio] || 0), nombre: f.Nombre || "", autor: f.Autor || "",
        hipotesis: f.Hipotesis || "", incrementos: _parse(f.IncrementosJSON),
        overrides: {}, medidas: {}, alojamiento: {}, personal: {},
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
    });
    localStorage.setItem(LS_VS, JSON.stringify(out));
    _conectado = true;
    return out;
  }

  // Guarda (upsert) la fila de metadatos de una versión desde localStorage.
  async function pushVersion(versionId) {
    if (!_conectado) { console.warn("pushVersion: sin conexión, no se guarda"); return; }
    try {
      var reg = (_parse(localStorage.getItem(LS_VS)))[versionId]; if (!reg) return;
      var cols = await getCols(C.listaVersiones), fAnio = interno(cols, "Año");
      var fields = { Title: versionId, Nombre: reg.nombre || "", Autor: reg.autor || "",
        Hipotesis: reg.hipotesis || "", IncrementosJSON: JSON.stringify(reg.incrementos || {}),
        Creada: reg.creada || new Date().toISOString(), Modificada: reg.modificada || new Date().toISOString() };
      fields[fAnio] = Number(reg.anio) || 0;
      var id = _idByTitle[C.listaVersiones] && _idByTitle[C.listaVersiones][versionId];
      if (id) await actualizarItem(C.listaVersiones, id, fields);
      else { var c = await crearItem(C.listaVersiones, fields); _idByTitle[C.listaVersiones][versionId] = c.id; }
    } catch (e) { console.error("Error guardando la versión en SharePoint:", e); }
  }

  // Guarda (upsert) la fila de un hotel de una versión desde localStorage.
  async function pushHotel(versionId, hotelId) {
    if (!_conectado) { console.warn("pushHotel: sin conexión, no se guarda"); return; }
    try {
      var reg = (_parse(localStorage.getItem(LS_VS)))[versionId]; if (!reg) return;
      var hid = String(hotelId), title = versionId + "_" + hid;
      var fields = { Title: title, VersionId: versionId, HotelId: Number(hotelId) || 0,
        OverridesJSON:   JSON.stringify((reg.overrides   && reg.overrides[hid])   || {}),
        MedidasJSON:     JSON.stringify((reg.medidas     && reg.medidas[hid])     || {}),
        AlojamientoJSON: JSON.stringify((reg.alojamiento && reg.alojamiento[hid]) || {}),
        PersonalJSON:    JSON.stringify((reg.personal    && reg.personal[hid])    || {}) };
      var id = _idByTitle[C.listaHoteles] && _idByTitle[C.listaHoteles][title];
      if (id) await actualizarItem(C.listaHoteles, id, fields);
      else { var c = await crearItem(C.listaHoteles, fields); _idByTitle[C.listaHoteles][title] = c.id; }
    } catch (e) { console.error("Error guardando el hotel en SharePoint:", e); }
  }

  // Sincroniza una vez por sesión al abrir cualquier pantalla en GitHub Pages: hace
  // login, descarga las versiones a localStorage y recarga (para que la página se
  // pinte con los datos frescos). En local no hace nada (modo localStorage).
  async function autoSync() {
    if (location.hostname.indexOf("github.io") === -1) return;   // modo local
    var primera = !sessionStorage.getItem("shg_synced");
    if (primera) document.documentElement.style.visibility = "hidden";  // evita el parpadeo
    try {
      await init();
      if (!cuenta()) { login(); return; }                        // redirige a login corporativo
      // SIEMPRE se descarga: así queda establecida la conexión (_conectado) y el
      // índice de ítems, para que los guardados funcionen también tras la recarga.
      await cargarTodo();
      if (primera) { sessionStorage.setItem("shg_synced", "1"); location.reload(); return; }
    } catch (e) {
      console.warn("Nube no disponible, modo local:", e);
      sessionStorage.setItem("shg_synced", "1");
    }
    document.documentElement.style.visibility = "";
  }
  function conectado() { return _conectado; }

  window.SHG = {
    init: init, login: login, logout: logout, cuenta: cuenta,
    graphCall: graphCall, getSiteId: getSiteId, getListId: getListId,
    getCols: getCols, interno: interno,
    leerItems: leerItems, crearItem: crearItem, actualizarItem: actualizarItem, borrarItem: borrarItem,
    cargarTodo: cargarTodo, pushVersion: pushVersion, pushHotel: pushHotel,
    autoSync: autoSync, conectado: conectado,
  };
})();
