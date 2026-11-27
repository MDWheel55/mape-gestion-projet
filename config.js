// Identifiants et emplacement des données. Rien de secret ici :
// clientId et tenantId sont publics par conception.
const CONFIG = {
  clientId: "37deec83-c9fc-4231-af2e-68613215ea05",  // « Mapé — gestion de projet »
  tenantId: "29b27da6-edf1-4191-a000-fe674cd08b36",  // ConseilMD

  hostname: "conseilmd.sharepoint.com",
  sitePath: "/sites/Mape",

  // Nom AFFICHÉ des listes dans SharePoint.
  lists: {
    orientations: "01-Orientations",
    projets:      "02-Projets",
    taches:       "03-Taches",
    carte:        "04-Carte"
  }
};
