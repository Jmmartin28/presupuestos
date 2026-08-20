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

  window.SHG = {
    init: init, login: login, logout: logout, cuenta: cuenta,
    graphCall: graphCall, getSiteId: getSiteId, getListId: getListId,
    getCols: getCols, interno: interno,
    leerItems: leerItems, crearItem: crearItem, actualizarItem: actualizarItem, borrarItem: borrarItem,
  };
})();
