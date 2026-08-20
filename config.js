/* Configuración de la app de Presupuestos SHG (Azure AD + SharePoint).
   El clientId y el tenant NO son secretos (app pública de navegador); se reutiliza
   el registro de Azure de Daily Cost. */
window.SHG_CONFIG = {
  clientId:     "b8b460c9-0535-4974-be26-d059b19c9386",
  tenant:       "a5839d11-3ae4-4e46-a6cb-db2abb961dfc",
  redirectUri:  "https://jmmartin28.github.io/presupuestos/",
  scopes:       ["https://graph.microsoft.com/Sites.ReadWrite.All",
                 "https://graph.microsoft.com/User.Read"],
  // Sitio de SharePoint (formato Graph: host:/sites/nombre:)
  site:         "sohohoteles.sharepoint.com:/sites/controldegestion:",
  listaVersiones: "Presupuesto_Versiones",
  listaHoteles:   "Presupuesto_Version_Hoteles",
};
