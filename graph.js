/* Socle réutilisable : connexion Entra ID (MSAL) + accès aux listes SharePoint via Microsoft Graph.
   Particularité de ce site : les listes ont été créées par import Excel, donc leurs noms internes
   sont « field_1, field_2… » alors que les noms affichés sont lisibles. On construit la
   correspondance au démarrage et on ne manipule plus que les noms affichés dans le reste du code. */
const SP = (function () {
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["Sites.ReadWrite.All"];   // passer à ["Sites.Selected"] une fois la permission accordée

  /* Adresse de retour = le dossier de la page, pas la racine du domaine : sur GitHub Pages,
     l'app vit dans https://<compte>.github.io/<dépôt>/, et location.origin + "/" renvoyait
     à la racine, hors de l'app. Les deux adresses ci-dessous sont à inscrire dans Entra. */
  const BASE = location.origin + location.pathname.replace(/[^/]*$/, "");
  const RETOUR_FENETRE = BASE + "blank.html";

  /* Dans un onglet « Site web » de Teams, la page tourne dans un cadre : Microsoft refuse
     d'afficher sa page de connexion dans un cadre. On passe alors par une fenêtre
     surgissante, ouverte depuis un clic (« Se connecter »), sinon le navigateur la bloque. */
  const DANS_CADRE = (function () { try { return window.self !== window.top; } catch (e) { return true; } })();

  const msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
      redirectUri: BASE
    },
    // localStorage : la session survit quand Teams recharge l'onglet (changement de canal).
    cache: { cacheLocation: "localStorage" }
  });

  let account = null;
  const ids = {};        // siteId + id de chaque liste
  const champs = {};     // par liste : { nomAffiché : nomInterne }

  /* Flux par REDIRECTION, pas par fenêtre surgissante : la page part chez Microsoft et
     revient sur elle-même. Aucun bloqueur de pop-ups ne peut s'y opposer, et ça fonctionne
     sur mobile. handleRedirectPromise() doit être appelé à chaque chargement, avant tout
     le reste : c'est lui qui ramasse le code d'autorisation présent dans l'adresse. */
  async function connexion(interactif) {
    await msalApp.initialize();

    const retour = await msalApp.handleRedirectPromise();
    if (retour && retour.account) {
      account = retour.account;
      msalApp.setActiveAccount(account);
      return account;
    }

    const dejaLa = msalApp.getAllAccounts();
    if (dejaLa.length) {
      account = dejaLa[0];
      msalApp.setActiveAccount(account);
      return account;
    }

    // Une authentification interrompue laisse un drapeau « interaction en cours » dans la
    // mémoire de session ; MSAL refuse alors toute nouvelle tentative. On le balaie ici,
    // mais seulement après handleRedirectPromise, qui a besoin de la trace d'un vrai retour.
    purgerInteraction();
    if (DANS_CADRE) {
      // Premier chargement dans Teams : on attend le clic sur « Se connecter ».
      if (!interactif) throw new Error("clique sur « Se connecter »");
      const r = await msalApp.loginPopup({ scopes: SCOPES, redirectUri: RETOUR_FENETRE });
      account = r.account;
      msalApp.setActiveAccount(account);
      return account;
    }
    await msalApp.loginRedirect({ scopes: SCOPES });
    // La page est remplacée par celle de Microsoft : on ne va pas plus loin.
    return new Promise(function () {});
  }

  function purgerInteraction() {
    try {
      for (let i = sessionStorage.length - 1; i >= 0; i--) {
        const k = sessionStorage.key(i);
        if (k && k.indexOf("interaction.status") !== -1) sessionStorage.removeItem(k);
      }
    } catch (e) { /* stockage indisponible : rien à purger */ }
  }
  function utilisateur() { return account ? (account.name || account.username) : ""; }

  let redirectionLancee = false;
  async function jeton() {
    try {
      return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
    } catch (e) {
      // Plusieurs appels Graph tournent en parallèle : sans ce garde-fou, chacun déclencherait
      // sa propre redirection et MSAL renverrait « interaction_in_progress ».
      if (DANS_CADRE) {
        return (await msalApp.acquireTokenPopup({ scopes: SCOPES, account, redirectUri: RETOUR_FENETRE })).accessToken;
      }
      if (redirectionLancee) return new Promise(function () {});
      redirectionLancee = true;
      purgerInteraction();
      await msalApp.acquireTokenRedirect({ scopes: SCOPES, account });
      return new Promise(function () {});
    }
  }

  async function appel(chemin, options = {}) {
    let url = chemin.startsWith("http") ? chemin : GRAPH + chemin;
    const tout = [];
    while (url) {
      const t = await jeton();
      const res = await fetch(url, {
        ...options,
        headers: { Authorization: "Bearer " + t, "Content-Type": "application/json", ...(options.headers || {}) }
      });
      if (res.status === 204) return null;
      const data = await res.json();
      if (!res.ok) throw new Error((data.error && data.error.message) || res.status);
      if (!data.value) return data;
      tout.push(...data.value);
      url = data["@odata.nextLink"] || null;
    }
    return tout;
  }

  async function siteId() {
    if (!ids.site) {
      const s = await appel("/sites/" + CONFIG.hostname + ":" + CONFIG.sitePath);
      ids.site = s.id;
    }
    return ids.site;
  }

  async function listeId(cle) {
    const nom = CONFIG.lists[cle];
    if (!ids[cle]) {
      const s = await siteId();
      const listes = await appel("/sites/" + s + "/lists?$select=id,displayName");
      const l = listes.find(x => x.displayName === nom);
      if (!l) throw new Error('Liste "' + nom + '" introuvable sur le site');
      ids[cle] = l.id;
    }
    return ids[cle];
  }

  /* Correspondance nom affiché → nom interne, lue une fois par liste.
     Attention : SharePoint expose plusieurs colonnes sous le même nom affiché. « Title » existe
     en trois exemplaires — Title (modifiable), LinkTitle et LinkTitleNoMenu (calculées, en
     lecture seule). Une correspondance naïve garde la dernière vue et l'écriture est refusée
     avec « Field 'LinkTitle' is read-only ». On tient donc deux tables : tout pour la lecture,
     les colonnes réellement modifiables pour l'écriture. */
  async function carte(cle) {
    if (!champs[cle]) {
      const s = await siteId(), l = await listeId(cle);
      const cols = await appel("/sites/" + s + "/lists/" + l +
        "/columns?$select=name,displayName,readOnly,hidden");
      const lecture = {}, ecriture = {};
      cols.forEach(c => {
        if (!(c.displayName in lecture)) lecture[c.displayName] = c.name;
        if (!c.readOnly && !c.hidden && !(c.displayName in ecriture)) ecriture[c.displayName] = c.name;
      });
      // Une colonne modifiable prime toujours pour la lecture aussi : c'est celle qui porte la valeur.
      Object.keys(ecriture).forEach(aff => { lecture[aff] = ecriture[aff]; });
      champs[cle] = { lecture, ecriture };
    }
    return champs[cle];
  }

  async function getItems(cle) {
    const s = await siteId(), l = await listeId(cle), m = (await carte(cle)).lecture;
    const inverse = {};
    Object.keys(m).forEach(aff => { inverse[m[aff]] = aff; });
    const items = await appel("/sites/" + s + "/lists/" + l + "/items?$expand=fields&$top=999");
    return items.map(it => {
      const o = { spId: it.id };
      Object.keys(it.fields || {}).forEach(interne => {
        const aff = inverse[interne];
        if (aff) o[aff] = it.fields[interne];
      });
      return o;
    });
  }

  // Traduit un objet { nomAffiché: valeur } en { nomInterne: valeur }, en ignorant l'inconnu
  async function traduire(cle, valeurs) {
    const m = (await carte(cle)).ecriture, out = {};
    Object.keys(valeurs).forEach(aff => {
      const interne = m[aff];
      if (interne) out[interne] = valeurs[aff];
    });
    return out;
  }

  async function addItem(cle, valeurs) {
    const s = await siteId(), l = await listeId(cle);
    return appel("/sites/" + s + "/lists/" + l + "/items",
      { method: "POST", body: JSON.stringify({ fields: await traduire(cle, valeurs) }) });
  }
  async function updateItem(cle, spId, valeurs) {
    const s = await siteId(), l = await listeId(cle);
    return appel("/sites/" + s + "/lists/" + l + "/items/" + spId + "/fields",
      { method: "PATCH", body: JSON.stringify(await traduire(cle, valeurs)) });
  }
  async function deleteItem(cle, spId) {
    const s = await siteId(), l = await listeId(cle);
    return appel("/sites/" + s + "/lists/" + l + "/items/" + spId, { method: "DELETE" });
  }

  // Dates : SharePoint renvoie un horodatage ISO ; on ne garde que le jour.
  // À l'écriture on vise midi UTC pour qu'aucun fuseau ne fasse basculer la veille.
  function jour(v) { return v ? String(v).slice(0, 10) : ""; }
  function iso(v) { return v ? v + "T12:00:00Z" : null; }
  function liste(v) { return v ? String(v).split(";").map(x => x.trim()).filter(Boolean) : []; }

  return { connexion, utilisateur, getItems, addItem, updateItem, deleteItem, jour, iso, liste };
})();
