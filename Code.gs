// ============================================
// CONSTANTES
// ============================================
// Conversion SCU ↔ cSCU : 1 SCU = 100 cSCU
const CSCU_PER_SCU = 100;

// ============================================
// GESTION DES UTILISATEURS
// ============================================

function authenticateUser(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('USERS');
    if (!usersSheet) return { success: false, error: 'Feuille USERS introuvable' };
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username && data[i][1] === password) {
        return {
          success: true,
          user: {
            username: data[i][0],
            displayName: data[i][2],
            reglementAccepte: data[i][3] === true || data[i][3] === 'TRUE',
            role: Number(data[i][4]) || 0,
            discordId: String(data[i][5] || '')   // colonne F : discord_id
          }
        };
      }
    }
    return { success: false, error: 'Identifiants incorrects' };
  } catch (e) {
    Logger.log('Erreur auth: ' + e);
    return { success: false, error: 'Erreur serveur' };
  }
}

// ============================================
// GESTION DES ITEMS (MARKETPLACE)
// ============================================

function getAllItems() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ITEMS');
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const range = sheet.getRange(2, 1, lastRow - 1, 14);
    const data = range.getValues();
    const result = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) {
        result.push({
          ID: String(data[i][0]), Date: _fmtDate(data[i][1]),
          Proprietaire: String(data[i][2]), Vendeur: String(data[i][3]),
          Item: String(data[i][4]), Description: String(data[i][5]),
          Prix: Number(data[i][6]), Quantite: Number(data[i][7]),
          PrixUnite: Boolean(data[i][8]), Categorie: String(data[i][9]),
          ImageURL: String(data[i][10] || ''), Statut: String(data[i][11]),
          Type: String(data[i][12] || 'Vente'), Qualite: Number(data[i][13]) || 500
        });
      }
    }
    return result;
  } catch (e) {
    Logger.log('Erreur getAllItems: ' + e);
    return [];
  }
}

function getUserItems(username) {
  return getAllItems().filter(item => item.Proprietaire === username);
}

function addItemWithOwner(itemData, username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ITEMS');
    if (!sheet) throw new Error('Feuille ITEMS introuvable');
    const data = sheet.getDataRange().getValues();
    let nextId = 1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) nextId = Math.max(nextId, parseInt(data[i][0]) + 1);
    }
    const date = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
    sheet.appendRow([nextId, date, username, itemData.vendeur, itemData.item,
      itemData.description, itemData.prix, itemData.quantite || 1, itemData.prixUnite || false,
      itemData.categorie, itemData.imageUrl || '', 'Disponible', itemData.type || 'Vente',
      itemData.qualite !== undefined ? Number(itemData.qualite) : 500]);
    return { success: true, id: nextId };
  } catch (error) {
    Logger.log('Erreur addItemWithOwner: ' + error);
    throw error;
  }
}

function updateItemStatus(itemId, newStatus, username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ITEMS');
    if (!sheet) return { success: false, error: 'Feuille ITEMS introuvable' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == itemId) {
        if (data[i][2] !== username) return { success: false, error: 'Vous ne pouvez modifier que vos propres annonces' };
        sheet.getRange(i + 1, 12).setValue(newStatus);
        return { success: true };
      }
    }
    return { success: false, error: 'Item non trouvé' };
  } catch (e) {
    Logger.log('Erreur updateItemStatus: ' + e);
    return { success: false, error: 'Erreur serveur' };
  }
}

function updateItem(itemData, username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ITEMS');
    if (!sheet) return { success: false, error: 'Feuille ITEMS introuvable' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == itemData.itemId) {
        if (data[i][2] !== username) return { success: false, error: 'Vous ne pouvez modifier que vos propres annonces' };
        const rowNum = i + 1;
        sheet.getRange(rowNum, 5).setValue(itemData.item);
        sheet.getRange(rowNum, 6).setValue(itemData.description);
        sheet.getRange(rowNum, 7).setValue(itemData.prix);
        sheet.getRange(rowNum, 8).setValue(itemData.quantite);
        sheet.getRange(rowNum, 9).setValue(itemData.prixUnite || false);
        sheet.getRange(rowNum, 10).setValue(itemData.categorie);
        sheet.getRange(rowNum, 11).setValue(itemData.imageUrl || '');
        sheet.getRange(rowNum, 13).setValue(itemData.type || 'Vente');
        sheet.getRange(rowNum, 14).setValue(itemData.qualite !== undefined ? Number(itemData.qualite) : 500);
        return { success: true };
      }
    }
    return { success: false, error: 'Item non trouvé' };
  } catch (e) {
    Logger.log('Erreur updateItem: ' + e);
    return { success: false, error: 'Erreur serveur' };
  }
}

function deleteItem(itemId, username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ITEMS');
    if (!sheet) return { success: false, error: 'Feuille ITEMS introuvable' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == itemId) {
        if (data[i][2] !== username) return { success: false, error: 'Vous ne pouvez supprimer que vos propres annonces' };
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Item non trouvé' };
  } catch (e) {
    Logger.log('Erreur deleteItem: ' + e);
    return { success: false, error: 'Erreur serveur' };
  }
}

function accepterReglementUser(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('USERS');
    if (!usersSheet) return { success: false, error: 'Feuille USERS introuvable' };
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        usersSheet.getRange(i + 1, 4).setValue(true);
        return { success: true };
      }
    }
    return { success: false, error: 'Utilisateur non trouvé' };
  } catch (e) {
    return { success: false, error: 'Erreur serveur' };
  }
}

// ============================================
// AUTHENTIFICATION VENDEUR (SESSIONS)
// ============================================

function loginVendeur(username, password) {
  try {
    const auth = authenticateUser(username, password);
    if (!auth.success) return { success: false, message: auth.error || 'Identifiants incorrects.' };
    _cleanExpiredTokens();
    const token = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      username + Date.now() + Math.random()
    ).map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('').substring(0, 32);
    const expires = Date.now() + 8 * 60 * 60 * 1000;
    PropertiesService.getScriptProperties().setProperty('token_' + token,
      JSON.stringify({ username: auth.user.username, displayName: auth.user.displayName,
        reglementAccepte: auth.user.reglementAccepte, role: auth.user.role || 0, expires: expires }));
    return { success: true, token: token };
  } catch (e) {
    Logger.log('Erreur loginVendeur: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function validateToken(token) {
  if (!token) return null;
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('token_' + token);
    if (!raw) return null;
    const session = JSON.parse(raw);
    if (Date.now() > session.expires) {
      PropertiesService.getScriptProperties().deleteProperty('token_' + token);
      return null;
    }
    return session;
  } catch (e) { return null; }
}

function _cleanExpiredTokens() {
  try {
    const props = PropertiesService.getScriptProperties().getProperties();
    const now = Date.now();
    Object.keys(props).forEach(key => {
      if (key.startsWith('token_')) {
        try {
          const s = JSON.parse(props[key]);
          if (now > s.expires) PropertiesService.getScriptProperties().deleteProperty(key);
        } catch (e) {}
      }
    });
  } catch (e) {}
}

// ============================================
// DONNÉES ESPACE VENDEUR
// ============================================

function getVendeurData(token) {
  try {
    const session = validateToken(token);
    if (!session) return { error: 'Session expirée. Veuillez vous reconnecter.' };
    const username = session.username;
    const allItems = getAllItems();
    const mine = allItems.filter(i => i.Proprietaire === username);
    const listings = mine.map(i => ({
      id: String(i.ID), item: i.Item, desc: i.Description, prix: i.Prix,
      quantite: i.Quantite, prixUnite: i.PrixUnite, categorie: i.Categorie,
      imageUrl: i.ImageURL, statut: i.Statut, type: i.Type, qualite: i.Qualite
    }));
    let reglementAccepte = false;
    try {
      const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
      if (usersSheet) {
        const usersData = usersSheet.getDataRange().getValues();
        for (let i = 1; i < usersData.length; i++) {
          if (usersData[i][0] === username) {
            reglementAccepte = usersData[i][3] === true || usersData[i][3] === 'TRUE';
            break;
          }
        }
      }
    } catch (e) {}
    return {
      vendorName: session.displayName || username, username: username,
      reglementAccepte: reglementAccepte, listings: listings,
      stats: {
        total: mine.length, dispo: mine.filter(i => i.Statut === 'Disponible').length,
        reserve: mine.filter(i => i.Statut === 'Réservé').length, vendu: mine.filter(i => i.Statut === 'Vendu').length
      }
    };
  } catch (e) {
    Logger.log('Erreur getVendeurData: ' + e);
    return { error: 'Erreur serveur.' };
  }
}

function saveListing(payload, token) {
  try {
    const session = validateToken(token);
    if (!session) return { success: false, message: 'Session expirée.' };
    const username = session.username;
    const itemData = {
      vendeur: session.displayName || username, item: payload.item,
      description: payload.desc || '', prix: Number(payload.prix) || 0,
      quantite: Number(payload.quantite) || 1, prixUnite: payload.prixUnite || false,
      categorie: payload.categorie || '', imageUrl: payload.imageUrl || '',
      type: payload.type || 'Vente', qualite: payload.qualite !== undefined ? Number(payload.qualite) : 500
    };
    if (payload.id) return updateItem(Object.assign({ itemId: payload.id }, itemData), username);
    else return addItemWithOwner(itemData, username);
  } catch (e) {
    Logger.log('Erreur saveListing: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function updateListingStatus(id, statut, token) {
  try {
    const session = validateToken(token);
    if (!session) return { success: false, message: 'Session expirée.' };
    return updateItemStatus(id, statut, session.username);
  } catch (e) {
    return { success: false, message: 'Erreur serveur.' };
  }
}

function deleteListing(id, token) {
  try {
    const session = validateToken(token);
    if (!session) return { success: false, message: 'Session expirée.' };
    return deleteItem(id, session.username);
  } catch (e) {
    return { success: false, message: 'Erreur serveur.' };
  }
}

function logoutVendeur(token) {
  try {
    if (token) PropertiesService.getScriptProperties().deleteProperty('token_' + token);
    return { success: true };
  } catch (e) { return { success: true }; }
}

// ============================================
// HELPERS INTERNES
// ============================================

function _getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function _nextId(sheet) {
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) max = Math.max(max, parseInt(data[i][0]) || 0);
  }
  return max + 1;
}

// Formate une valeur date (Date objet ou string) en dd/MM/yyyy HH:mm heure France
function _fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Paris', 'dd/MM/yyyy HH:mm');
  return String(v);
}

function _fmtDay(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Europe/Paris', 'dd/MM/yyyy');
  return String(v).slice(0, 10);
}

function _logStock(username, details) {
  try {
    const sheet = _getSheet('STOCK_LOG');
    if (!sheet) return;
    const id = _nextId(sheet);
    const date = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
    sheet.appendRow([id, date, username, details]);
  } catch (e) {}
}

function _safeParse(str, fallback) {
  try {
    if (!str || str === '') return fallback;
    return JSON.parse(String(str));
  } catch (e) { return fallback; }
}

// ============================================
// STOCK
// STOCK sheet: ID(0), Categorie(1), Item(2), Quantite(3), Qualite(4), Unite(5), ImageURL(6),
//              Reserve(7), Seuil1(8), Seuil2(9), Seuil3(10), Actif(11)
// ============================================

function _isActif(val) {
  // Returns true if active (default: true when empty)
  return val !== false && val !== 'FALSE' && val !== 0;
}

function getStock() {
  try {
    const sheet = _getSheet('STOCK');
    if (!sheet || sheet.getLastRow() <= 1) return { categories: [] };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();

    // Cible + catégorie par item depuis MATERIALS_CONFIG
    const cibleMap = {};
    const catUnitMap = _getCategoryUnitMap();
    try {
      const cfgSheet = _getSheet('MATERIALS_CONFIG');
      if (cfgSheet && cfgSheet.getLastRow() > 1) {
        cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 7).getValues().forEach(function(r) {
          if (r[0]) cibleMap[String(r[0]).toLowerCase()] = Number(r[5]) || 0;
        });
      }
    } catch(e) {}

    const catMap = {};
    data.forEach(function(row) {
      if (!row[0]) return;
      const cat    = String(row[1] || 'Autres');
      const name   = String(row[2]);
      const actif  = _isActif(row[11]);
      const seuil1 = Number(row[8]) || 0;
      const seuil2 = Number(row[9]) || 0;
      const seuil3 = Number(row[10]) || 0;
      const unite  = catUnitMap[cat.toLowerCase()] || 'cSCU';

      if (!catMap[cat]) catMap[cat] = {};
      if (!catMap[cat][name]) {
        catMap[cat][name] = {
          name: name, imageUrl: String(row[6] || ''),
          actif: actif, seuil1: seuil1, seuil2: seuil2, seuil3: seuil3,
          unite: unite, entries: []
        };
      }
      catMap[cat][name].entries.push({
        id: String(row[0]), quantite: Number(row[3]) || 0,
        qualite: Number(row[4]) || 0, unite: unite,
        reserve: Number(row[7]) || 0
      });
      // Accumulate thresholds (use max across all entries of same item)
      if (seuil1 > catMap[cat][name].seuil1) catMap[cat][name].seuil1 = seuil1;
      if (seuil2 > catMap[cat][name].seuil2) catMap[cat][name].seuil2 = seuil2;
      if (seuil3 > catMap[cat][name].seuil3) catMap[cat][name].seuil3 = seuil3;
      if (!actif) catMap[cat][name].actif = false; // any inactive entry marks item as inactive
    });

    // Sort categories alphabetically (Gemmes < Matériaux)
    const categories = Object.keys(catMap).sort(function(a, b) {
      return a.localeCompare(b, 'fr');
    }).map(function(cat) {
      const items = Object.values(catMap[cat]).map(function(item) {
        // Sort entries by quality DESC
        item.entries.sort(function(a, b) { return b.qualite - a.qualite; });
        const totalQty  = item.entries.reduce(function(s, e) { return s + e.quantite; }, 0);
        const totalAvail = item.entries.reduce(function(s, e) { return s + Math.max(0, e.quantite - e.reserve); }, 0);
        const totalReserve = item.entries.reduce(function(s, e) { return s + (e.reserve || 0); }, 0);
        const bestEntry = item.entries.length > 0 ? item.entries[0] : null;
        const minorTotal = item.entries.slice(1).reduce(function(s, e) { return s + e.quantite; }, 0);

        // Alert level: 3=critique, 2=modéré, 1=attention, 0=aucun
        // Items at 0 always get level 3 so they stay visible and sorted to top
        let alertLevel = 0;
        if (totalAvail <= 0) alertLevel = 3;
        else if (item.seuil3 > 0 && totalAvail <= item.seuil3) alertLevel = 3;
        else if (item.seuil2 > 0 && totalAvail <= item.seuil2) alertLevel = 2;
        else if (item.seuil1 > 0 && totalAvail <= item.seuil1) alertLevel = 1;

        return {
          name: item.name, imageUrl: item.imageUrl, actif: item.actif,
          bestEntry: bestEntry, minorTotal: minorTotal,
          totalQty: totalQty, totalAvail: totalAvail, totalReserve: totalReserve,
          alertLevel: alertLevel, isEmpty: totalAvail <= 0,
          seuil1: item.seuil1, seuil2: item.seuil2, seuil3: item.seuil3,
          cible: cibleMap[item.name.toLowerCase()] || 0,
          unite: item.unite,
          entries: item.entries
        };
      });
      // Sort: empty items first, then by alert level descending, then alphabetically
      items.sort(function(a, b) {
        if (a.isEmpty && !b.isEmpty) return -1;
        if (!a.isEmpty && b.isEmpty) return 1;
        if (b.alertLevel !== a.alertLevel) return b.alertLevel - a.alertLevel;
        return a.name.localeCompare(b.name, 'fr');
      });
      return { name: cat, items: items };
    });
    return { categories: categories };
  } catch (e) {
    Logger.log('Erreur getStock: ' + e);
    return { categories: [] };
  }
}

function getStockAdmin(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { error: 'Accès refusé.' };
    const sheet = _getSheet('STOCK');
    if (!sheet || sheet.getLastRow() <= 1) return { items: [], log: [] };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
    const adminCibleMap = {};
    const catUnitMap = _getCategoryUnitMap();
    try {
      const cfgSheet = _getSheet('MATERIALS_CONFIG');
      if (cfgSheet && cfgSheet.getLastRow() > 1) {
        cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 7).getValues().forEach(function(r) {
          if (r[0]) adminCibleMap[String(r[0]).toLowerCase()] = Number(r[5]) || 0;
        });
      }
    } catch(e) {}
    const items = data.filter(function(r) { return r[0]; }).map(function(r) {
      const cat = String(r[1] || '');
      return {
        id: String(r[0]), categorie: cat, item: String(r[2] || ''),
        quantite: Number(r[3]) || 0, qualite: Number(r[4]) || 0,
        unite: catUnitMap[cat.toLowerCase()] || 'cSCU',
        imageUrl: String(r[6] || ''),
        reserve: Number(r[7]) || 0,
        seuil1: Number(r[8]) || 0, seuil2: Number(r[9]) || 0, seuil3: Number(r[10]) || 0,
        actif: _isActif(r[11]),
        cible: adminCibleMap[String(r[2] || '').toLowerCase()] || 0
      };
    });
    const log = getStockLog(token).logs || [];
    return { items: items, log: log };
  } catch (e) {
    return { error: 'Erreur serveur.' };
  }
}

function addStockItem(payload, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    const sheet = _getSheet('STOCK');
    if (!sheet) return { success: false, message: 'Feuille STOCK introuvable.' };
    const id = _nextId(sheet);
    sheet.appendRow([
      id, payload.categorie || '', payload.item || '',
      Number(payload.quantite) || 0, Number(payload.qualite) || 0,
      payload.unite || 'unité', payload.imageUrl || '', 0,
      Number(payload.seuil1) || 0, Number(payload.seuil2) || 0, Number(payload.seuil3) || 0,
      payload.actif !== false
    ]);
    _logStock(session.displayName || session.username,
      'Ajout : ' + payload.item + ' x' + payload.quantite + ' (Q:' + payload.qualite + ')');
    return { success: true, id: id };
  } catch (e) {
    Logger.log('Erreur addStockItem: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function updateStockItem(payload, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    const sheet = _getSheet('STOCK');
    if (!sheet) return { success: false, message: 'Feuille STOCK introuvable.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.id)) {
        const r = i + 1;
        sheet.getRange(r, 2).setValue(payload.categorie !== undefined ? payload.categorie : data[i][1]);
        sheet.getRange(r, 3).setValue(payload.item !== undefined ? payload.item : data[i][2]);
        sheet.getRange(r, 4).setValue(Number(payload.quantite) || 0);
        sheet.getRange(r, 5).setValue(Number(payload.qualite) || 0);
        sheet.getRange(r, 6).setValue(payload.unite !== undefined ? payload.unite : data[i][5]);
        sheet.getRange(r, 7).setValue(payload.imageUrl !== undefined ? payload.imageUrl : data[i][6]);
        sheet.getRange(r, 9).setValue(Number(payload.seuil1) || 0);
        sheet.getRange(r, 10).setValue(Number(payload.seuil2) || 0);
        sheet.getRange(r, 11).setValue(Number(payload.seuil3) || 0);
        sheet.getRange(r, 12).setValue(payload.actif !== false);
        _logStock(session.displayName || session.username,
          'Modif : ' + (payload.item || data[i][2]) + ' → x' + payload.quantite + ' (Q:' + payload.qualite + ')');
        return { success: true };
      }
    }
    return { success: false, message: 'Item non trouvé.' };
  } catch (e) {
    Logger.log('Erreur updateStockItem: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function deleteStockItem(id, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    const sheet = _getSheet('STOCK');
    if (!sheet) return { success: false, message: 'Feuille introuvable.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const qty = Number(data[i][3]) || 0;
        if (qty > 0) {
          // Qty still positive → reset to 0/500 so it stays visible as "to mine"
          sheet.getRange(i + 1, 4).setValue(0);
          sheet.getRange(i + 1, 5).setValue(500);
          sheet.getRange(i + 1, 8).setValue(0); // clear reservation
          _logStock(session.displayName || session.username,
            'Réinitialisation (à miner) : ' + data[i][2] + ' x' + data[i][3]);
          return { success: true, reset: true, message: 'Stock remis à zéro — conservé visible (à miner).' };
        }
        // Already at 0 → actually delete
        _logStock(session.displayName || session.username,
          'Suppression : ' + data[i][2]);
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Item non trouvé.' };
  } catch (e) {
    return { success: false, message: 'Erreur serveur.' };
  }
}

function getStockLog(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { error: 'Accès refusé.' };
    const sheet = _getSheet('STOCK_LOG');
    if (!sheet || sheet.getLastRow() <= 1) return { logs: [] };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    const logs = data.filter(function(r) { return r[0]; }).map(function(r) {
      return { id: String(r[0]), date: _fmtDate(r[1]), by: String(r[2]), details: String(r[3]) };
    }).reverse();
    return { logs: logs };
  } catch (e) { return { logs: [] }; }
}

// ============================================
// BLUEPRINT ITEMS & RECETTES
// CRAFT_ITEMS: ID(0), NomItem(1), Description(2), ImageURL(3)
// CRAFT_RECIPES: ID(0), CraftItemID(1), ResourceNom(2), QuantiteNecessaire(3)
// ============================================

// Returns {itemNameLower: factor} where factor=1 for 'unité' items, CSCU_PER_SCU for 'cSCU' items
function _getResourceUnitFactors() {
  const catUnitMap = _getCategoryUnitMap();
  const factors = {};
  // Primary: MATERIALS_CONFIG col7 — the category set via Catalogue UI (authoritative)
  const cfgSheet = _getSheet('MATERIALS_CONFIG');
  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 7).getValues().forEach(function(r) {
      if (!r[0]) return;
      const cat = String(r[6] || '');
      const unitType = catUnitMap[cat.toLowerCase()] || 'cSCU';
      factors[String(r[0]).toLowerCase()] = unitType === 'unité' ? 1 : CSCU_PER_SCU;
    });
  }
  // Fallback: STOCK col2 for items not listed in MATERIALS_CONFIG
  const stockSheet = _getSheet('STOCK');
  if (stockSheet && stockSheet.getLastRow() > 1) {
    stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 3).getValues().forEach(function(r) {
      if (!r[0] || !r[2]) return;
      const name = String(r[2]).toLowerCase();
      if (factors[name] !== undefined) return;
      const unitType = catUnitMap[String(r[1] || '').toLowerCase()] || 'cSCU';
      factors[name] = unitType === 'unité' ? 1 : CSCU_PER_SCU;
    });
  }
  return factors;
}

function _buildStockMap() {
  const stockSheet = _getSheet('STOCK');
  const map = {};
  if (stockSheet && stockSheet.getLastRow() > 1) {
    const d = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 12).getValues();
    d.forEach(function(r) {
      if (!r[0]) return;
      if (!_isActif(r[11])) return; // Skip disabled resources
      const k = String(r[2]).toLowerCase();
      map[k] = (map[k] || 0) + Math.max(0, (Number(r[3]) || 0) - (Number(r[7]) || 0));
    });
  }
  return map;
}

function getStockSummary(token) {
  const session = validateToken(token);
  if (!session) return { error: 'Connexion requise.' };
  return { summary: _buildStockMap() };
}

function _buildRecipesMap() {
  const sheet = _getSheet('CRAFT_RECIPES');
  const map = {};
  if (sheet && sheet.getLastRow() > 1) {
    const d = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    d.forEach(function(r) {
      if (!r[0]) return;
      const cid = String(r[1]);
      if (!map[cid]) map[cid] = [];
      map[cid].push({ id: String(r[0]), resource: String(r[2]), qty: Number(r[3]) || 0 });
    });
  }
  return map;
}

function _isWikiUUID(id) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(id));
}

function _getActiveWikiUuids() {
  const sheet = _getSheet('WIKI_ACTIVE');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .filter(function(r) { return r[0]; }).map(function(r) { return String(r[0]); });
}

function _getWikiIngredientsForCraft(uuid) {
  const sheet = _getSheet('WIKI_BLUEPRINTS');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1]) === String(uuid)) {
      return _safeParse(data[i][5], []).map(function(ing) {
        return { resource: String(ing.name || ''), qty: Number(ing.qty || 0) };
      });
    }
  }
  return [];
}

function getCraftItemsWithStatus() {
  try {
    const activeUuids = _getActiveWikiUuids();
    if (activeUuids.length === 0) return { items: [] };
    const wikiSheet = _getSheet('WIKI_BLUEPRINTS');
    if (!wikiSheet || wikiSheet.getLastRow() <= 1) return { items: [] };
    const stockMap = _buildStockMap();
    const unitFactors = _getResourceUnitFactors();
    const data = wikiSheet.getRange(2, 1, wikiSheet.getLastRow() - 1, 10).getValues();
    const items = [];
    data.forEach(function(r) {
      const uuid = String(r[1] || '');
      if (!uuid || activeUuids.indexOf(uuid) === -1) return;
      const ingredients = _safeParse(r[5], []);
      const recipe = ingredients.map(function(ing) {
        return { resource: String(ing.name || ''), qty: Number(ing.qty || 0) };
      });
      let craftable = recipe.length > 0;
      recipe.forEach(function(ing) {
        const factor = unitFactors[ing.resource.toLowerCase()] !== undefined ? unitFactors[ing.resource.toLowerCase()] : CSCU_PER_SCU;
        if (ing.qty > 0 && (stockMap[ing.resource.toLowerCase()] || 0) < ing.qty * factor) craftable = false;
      });
      items.push({
        id: uuid,
        nom: String(r[2] || ''),
        description: '',
        imageUrl: '',
        recipe: recipe,
        craftable: recipe.length === 0 ? false : craftable,
        craftTime: Number(r[4] || 0),
        type: String(r[9] || ''),
        webUrl: String(r[6] || '')
      });
    });
    return { items: items };
  } catch (e) {
    Logger.log('Erreur getCraftItemsWithStatus: ' + e);
    return { items: [] };
  }
}

function getCraftItemsAdmin(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { error: 'Accès refusé.' };
    const itemsSheet = _getSheet('CRAFT_ITEMS');
    if (!itemsSheet || itemsSheet.getLastRow() <= 1) return { items: [] };
    const data = itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 4).getValues();
    const recipesMap = _buildRecipesMap();
    const items = data.filter(function(r) { return r[1]; }).map(function(r) {
      return { id: String(r[0]), nom: String(r[1]), description: String(r[2] || ''),
               imageUrl: String(r[3] || ''), recipe: recipesMap[String(r[0])] || [] };
    });
    return { items: items };
  } catch (e) { return { items: [] }; }
}

function saveCraftItem(payload, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    const sheet = _getSheet('CRAFT_ITEMS');
    if (!sheet) return { success: false, message: 'Feuille CRAFT_ITEMS introuvable.' };
    let itemId = payload.id;
    if (itemId) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(itemId)) {
          sheet.getRange(i + 1, 2).setValue(payload.nom || '');
          sheet.getRange(i + 1, 3).setValue(payload.description || '');
          sheet.getRange(i + 1, 4).setValue(payload.imageUrl || '');
          break;
        }
      }
    } else {
      itemId = String(_nextId(sheet));
      sheet.appendRow([itemId, payload.nom || '', payload.description || '', payload.imageUrl || '']);
    }
    if (Array.isArray(payload.recipe)) _saveRecipes(itemId, payload.recipe);
    return { success: true, id: itemId };
  } catch (e) {
    Logger.log('Erreur saveCraftItem: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function _saveRecipes(craftItemId, recipes) {
  const sheet = _getSheet('CRAFT_RECIPES');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(craftItemId)) sheet.deleteRow(i + 1);
  }
  recipes.forEach(function(r) {
    if (!r.resource || !r.qty) return;
    sheet.appendRow([_nextId(sheet), craftItemId, r.resource, Number(r.qty) || 0]);
  });
}

function deleteCraftItem(id, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    const sheet = _getSheet('CRAFT_ITEMS');
    if (!sheet) return { success: false, message: 'Feuille introuvable.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        _saveRecipes(String(id), []);
        return { success: true };
      }
    }
    return { success: false, message: 'Item non trouvé.' };
  } catch (e) { return { success: false, message: 'Erreur serveur.' }; }
}

// ============================================
// BLUEPRINT REQUESTS
// CRAFT_REQUESTS: ID(0), Date(1), Username(2), DisplayName(3), CraftItemID(4), NomItem(5),
//   Quantite(6), Note(7), Statut(8), NoteAdmin(9), DateTraitement(10), TraitePar(11),
//   PrioriteQualite(12), ReservationDetails(13), AdminAccepte(14), UserAccepte(15),
//   Messages(16), CustomRecipe(17), QualitesChoisies(18)
// ============================================

function submitCraftRequest(payload, token) {
  try {
    const session = validateToken(token);
    if (!session) return { success: false, message: 'Connexion requise.' };
    const sheet = _getSheet('CRAFT_REQUESTS');
    if (!sheet) return { success: false, message: 'Feuille CRAFT_REQUESTS introuvable.' };
    const id   = _nextId(sheet);
    const date = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
    const customRecipe    = payload.customRecipe ? JSON.stringify(payload.customRecipe) : '';
    const qualitesChoisies = payload.qualitesChoisies ? JSON.stringify(payload.qualitesChoisies) : '';
    sheet.appendRow([
      id, date, session.username, session.displayName || session.username,
      payload.craftItemId || '', payload.nomItem || '',
      Number(payload.quantite) || 1, payload.note || '',
      'En attente', '', '', '',
      payload.prioriteQualite || 'DESC', '',
      false, false, '[]',
      customRecipe, qualitesChoisies
    ]);

    // [DISCORD BOT] Notifier les admins d'une nouvelle demande de craft
    _discordNotifyAdmins(
      '🔔 **Nouvelle demande de craft**\n' +
      'De : **' + (session.displayName || session.username) + '**\n' +
      'Item : **' + (payload.nomItem || payload.craftItemId || '?') + '**' +
      (payload.note ? '\nNote : ' + String(payload.note).substring(0, 200) : '')
    );

    return { success: true, id: id };
  } catch (e) {
    Logger.log('Erreur submitCraftRequest: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function sendCraftMessage(id, text, token) {
  try {
    const session = validateToken(token);
    if (!session) return { success: false, message: 'Connexion requise.' };
    const sheet = _getSheet('CRAFT_REQUESTS');
    if (!sheet) return { success: false, message: 'Feuille introuvable.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const isAdmin = (session.role || 0) >= 1;
        if (!isAdmin && String(data[i][2]) !== session.username) {
          return { success: false, message: 'Accès refusé.' };
        }
        const row = i + 1;
        const messages = _safeParse(data[i][16], []);
        const date = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
        messages.push({
          from: isAdmin ? 'admin' : 'user',
          displayName: session.displayName || session.username,
          text: String(text).substring(0, 500),
          date: date
        });
        sheet.getRange(row, 17).setValue(JSON.stringify(messages));
        // Move to "En discussion" if not already validated/refused
        const currentStatut = String(data[i][8]);
        if (currentStatut === 'En attente') {
          sheet.getRange(row, 9).setValue('En discussion');
        }
        // Notification Discord MP
        const nomItem = String(data[i][5]);
        if (isAdmin) {
          // Admin écrit → MP au demandeur
          _discordNotifyUser(String(data[i][2]),
            '💬 **Nouveau message du crafteur**\n' +
            'Concernant ta demande **' + nomItem + '** :\n> ' + String(text).substring(0, 400));
        } else {
          // Membre écrit → MP aux admins
          _discordNotifyAdmins(
            '💬 **Nouveau message de ' + (session.displayName || session.username) + '**\n' +
            'Demande : **' + nomItem + '**\n> ' + String(text).substring(0, 400));
        }
        return { success: true };
      }
    }
    return { success: false, message: 'Demande non trouvée.' };
  } catch (e) {
    Logger.log('Erreur sendCraftMessage: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function cancelCraftRequest(id, token) {
  try {
    const session = validateToken(token);
    if (!session) return { success: false, message: 'Connexion requise.' };
    const sheet = _getSheet('CRAFT_REQUESTS');
    if (!sheet) return { success: false, message: 'Feuille introuvable.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        // Seul le demandeur peut annuler sa propre demande (ou un admin)
        const isAdmin = (session.role || 0) >= 1;
        if (!isAdmin && String(data[i][2]) !== session.username) {
          return { success: false, message: 'Accès refusé.' };
        }
        const oldStatut = String(data[i][8]);
        if (oldStatut === 'Crafté') {
          return { success: false, message: 'Un blueprint déjà crafté ne peut pas être annulé.' };
        }
        const row = i + 1;
        sheet.getRange(row, 9).setValue('Annulé');
        const date = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
        sheet.getRange(row, 11).setValue(date);
        sheet.getRange(row, 12).setValue(session.displayName || session.username);
        // Libérer la réservation si la demande était Acceptée
        if (oldStatut === 'Accepté') {
          const customRecipe = _safeParse(data[i][17], null);
          _updateStockReservation(String(data[i][4]), Number(data[i][6]) || 1, -1, String(data[i][12] || 'DESC'), null, customRecipe);
          sheet.getRange(row, 14).setValue('');
        }
        return { success: true };
      }
    }
    return { success: false, message: 'Demande non trouvée.' };
  } catch (e) {
    Logger.log('Erreur cancelCraftRequest: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function acceptCraftRequest(id, token) {
  try {
    const session = validateToken(token);
    if (!session) return { success: false, message: 'Connexion requise.' };
    if ((session.role || 0) < 1) return { success: false, message: 'Seul le crafteur peut valider une demande.' };
    const sheet = _getSheet('CRAFT_REQUESTS');
    if (!sheet) return { success: false, message: 'Feuille introuvable.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const currentStatut = String(data[i][8]);
        if (currentStatut === 'Refusé' || currentStatut === 'Crafté') {
          return { success: false, message: 'Cette demande ne peut plus être modifiée.' };
        }
        const row = i + 1;
        sheet.getRange(row, 15).setValue(true); // adminAccepte
        sheet.getRange(row, 9).setValue('Accepté');
        const date = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
        sheet.getRange(row, 11).setValue(date);
        sheet.getRange(row, 12).setValue(session.displayName || session.username);
        const craftItemId      = String(data[i][4]);
        const quantite         = Number(data[i][6]) || 1;
        const priorite         = String(data[i][12] || 'DESC');
        const qualitesChoisies = _safeParse(data[i][18], null);
        const customRecipe     = _safeParse(data[i][17], null);
        const details = _updateStockReservation(craftItemId, quantite, 1, priorite, qualitesChoisies, customRecipe);
        sheet.getRange(row, 14).setValue(JSON.stringify(details));
        return { success: true };
      }
    }
    return { success: false, message: 'Demande non trouvée.' };
  } catch (e) {
    Logger.log('Erreur acceptCraftRequest: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function getMyCraftRequests(token) {
  try {
    const session = validateToken(token);
    if (!session) return { error: 'Connexion requise.' };
    const sheet = _getSheet('CRAFT_REQUESTS');
    if (!sheet || sheet.getLastRow() <= 1) return { requests: [] };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 19).getValues();
    const requests = data
      .filter(function(r) { return r[0] && String(r[2]) === session.username; })
      .map(function(r) {
        return {
          id:          String(r[0]),
          date:        _fmtDate(r[1]),
          nomItem:     String(r[5]),
          quantite:    Number(r[6]),
          note:        String(r[7]),
          statut:      String(r[8]),
          noteAdmin:   String(r[9] || ''),
          dateTraitement: _fmtDate(r[10] || ''),
          reservationDetails: _safeParse(r[13], []),
          adminAccepte: r[14] === true || r[14] === 'TRUE',
          userAccepte:  r[15] === true || r[15] === 'TRUE',
          messages:    _safeParse(r[16], []),
          customRecipe: _safeParse(r[17], null),
          qualitesChoisies: _safeParse(r[18], null)
        };
      }).reverse();
    return { requests: requests, displayName: session.displayName || session.username };
  } catch (e) {
    Logger.log('Erreur getMyCraftRequests: ' + e);
    return { requests: [] };
  }
}

function getCraftRequests(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { error: 'Accès refusé.' };
    const sheet = _getSheet('CRAFT_REQUESTS');
    if (!sheet || sheet.getLastRow() <= 1) return { requests: [] };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 19).getValues();
    const stockMap = _buildStockMap();
    const unitFactors = _getResourceUnitFactors();
    const requests = data.filter(function(r) { return r[0]; }).map(function(r) {
      const customRecipe = _safeParse(r[17], null);
      const quantite = Number(r[6]) || 1;
      // Compute feasibility for custom blueprints
      let feasible = null;
      if (customRecipe && Array.isArray(customRecipe)) {
        feasible = customRecipe.every(function(ing) {
          const factor = unitFactors[String(ing.resource).toLowerCase()] !== undefined ? unitFactors[String(ing.resource).toLowerCase()] : CSCU_PER_SCU;
          return (stockMap[String(ing.resource).toLowerCase()] || 0) >= (Number(ing.qty) || 0) * quantite * factor;
        });
      }
      return {
        id:           String(r[0]), date: _fmtDate(r[1]),
        username:     String(r[2]), displayName: String(r[3]),
        craftItemId:  String(r[4]), nomItem: String(r[5]),
        quantite:     quantite, note: String(r[7]), statut: String(r[8]),
        noteAdmin:    String(r[9] || ''), dateTraitement: String(r[10] || ''),
        traitePar:    String(r[11] || ''), prioriteQualite: String(r[12] || 'DESC'),
        reservationDetails: _safeParse(r[13], []),
        adminAccepte: r[14] === true || r[14] === 'TRUE',
        userAccepte:  r[15] === true || r[15] === 'TRUE',
        messages:     _safeParse(r[16], []),
        customRecipe: customRecipe, feasible: feasible,
        qualitesChoisies: _safeParse(r[18], null)
      };
    }).reverse();
    return { requests: requests };
  } catch (e) {
    Logger.log('Erreur getCraftRequests: ' + e);
    return { requests: [] };
  }
}

function updateCraftRequest(id, statut, noteAdmin, token) {
  // Only used for Refusé and Crafté (acceptance handled by acceptCraftRequest)
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    const sheet = _getSheet('CRAFT_REQUESTS');
    if (!sheet) return { success: false, message: 'Feuille introuvable.' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        const oldStatut = String(data[i][8]);
        const row = i + 1;
        sheet.getRange(row, 9).setValue(statut);
        if (noteAdmin !== null && noteAdmin !== undefined) {
          sheet.getRange(row, 10).setValue(noteAdmin);
        }
        if (statut !== 'En attente' && statut !== 'En discussion') {
          const date = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
          sheet.getRange(row, 11).setValue(date);
          sheet.getRange(row, 12).setValue(session.displayName || session.username);
        }
        // Release reservation if previously Accepté
        if (oldStatut === 'Accepté' && (statut === 'Refusé' || statut === 'En attente' || statut === 'En discussion')) {
          const customRecipe = _safeParse(data[i][17], null);
          _updateStockReservation(String(data[i][4]), Number(data[i][6]) || 1, -1, String(data[i][12] || 'DESC'), null, customRecipe);
          sheet.getRange(row, 14).setValue('');
          sheet.getRange(row, 15).setValue(false);
          sheet.getRange(row, 16).setValue(false);
        }
        if (statut === 'Crafté' && oldStatut === 'Accepté') {
          const customRecipe = _safeParse(data[i][17], null);
          _updateStockReservation(String(data[i][4]), Number(data[i][6]) || 1, -1, String(data[i][12] || 'DESC'), null, customRecipe);
          sheet.getRange(row, 14).setValue('');
        }

        // [DISCORD BOT] Notifier le demandeur du changement de statut
        if (statut !== oldStatut) {
          const nomItem    = String(data[i][5] || data[i][4] || '?');
          const demandeur  = String(data[i][2]);
          const statusEmoji = { 'Accepté': '✅', 'Refusé': '❌', 'Crafté': '🎉', 'En discussion': '💬', 'En attente': '⏳' };
          const emoji = statusEmoji[statut] || '🔔';
          _discordNotifyUser(demandeur,
            emoji + ' **Mise à jour de ta demande de craft**\n' +
            'Item : **' + nomItem + '**\n' +
            'Nouveau statut : **' + statut + '**' +
            (noteAdmin ? '\nNote : ' + String(noteAdmin).substring(0, 300) : '')
          );
        }

        return { success: true };
      }
    }
    return { success: false, message: 'Demande non trouvée.' };
  } catch (e) {
    Logger.log('Erreur updateCraftRequest: ' + e);
    return { success: false, message: 'Erreur serveur.' };
  }
}

function _updateStockReservation(craftItemId, quantite, direction, prioriteQualite, qualitesChoisies, customRecipe) {
  const allDetails = [];
  try {
    const stockSheet = _getSheet('STOCK');
    if (!stockSheet) return allDetails;
    const sd = stockSheet.getDataRange().getValues();
    const unitFactors = _getResourceUnitFactors();

    // Build recipe list
    let recipes = [];
    if (customRecipe && Array.isArray(customRecipe)) {
      recipes = customRecipe;
    } else if (_isWikiUUID(craftItemId)) {
      recipes = _getWikiIngredientsForCraft(craftItemId);
    } else if (craftItemId) {
      const recipesSheet = _getSheet('CRAFT_RECIPES');
      if (!recipesSheet) return allDetails;
      const rd = recipesSheet.getDataRange().getValues();
      rd.forEach(function(r) {
        if (r[0] && String(r[1]) === String(craftItemId)) {
          recipes.push({ resource: String(r[2]), qty: Number(r[3]) || 0 });
        }
      });
    }

    recipes.forEach(function(recipe) {
      const resName = String(recipe.resource);
      const factor = unitFactors[resName.toLowerCase()] !== undefined ? unitFactors[resName.toLowerCase()] : CSCU_PER_SCU;
      let remaining = Math.round((Number(recipe.qty) || 0) * quantite * factor * 10000) / 10000;

      // Determine sort priority and optional minimum quality for this resource
      let localPriority = prioriteQualite || 'DESC';
      let minQualityFilter = null;
      if (qualitesChoisies) {
        const qc = qualitesChoisies[resName.toLowerCase()];
        if (qc) {
          if (qc.type === 'best') localPriority = 'DESC';
          else if (qc.type === 'worst') localPriority = 'ASC';
          else if (qc.type !== 'any') {
            const minQ = parseInt(qc.type);
            if (!isNaN(minQ)) { minQualityFilter = minQ; localPriority = 'DESC'; }
          }
        }
      }

      const sortFn = localPriority === 'ASC'
        ? function(a, b) { return (Number(a.r[4]) || 0) - (Number(b.r[4]) || 0); }
        : function(a, b) { return (Number(b.r[4]) || 0) - (Number(a.r[4]) || 0); };

      const entries = sd.map(function(r, idx) { return { r: r, idx: idx }; })
        .filter(function(e) {
          if (!e.r[0]) return false;
          if (String(e.r[2]).toLowerCase() !== resName.toLowerCase()) return false;
          if (minQualityFilter !== null && (Number(e.r[4]) || 0) < minQualityFilter) return false;
          return true;
        })
        .sort(sortFn);

      entries.forEach(function(e) {
        if (remaining <= 0) return;
        const cur   = Number(e.r[7]) || 0;
        const apply = Math.min(remaining, Number(e.r[3]) || 0);
        const newVal = Math.max(0, cur + direction * apply);
        stockSheet.getRange(e.idx + 1, 8).setValue(newVal);
        if (direction > 0 && apply > 0) {
          allDetails.push({ resource: resName, qualite: Number(e.r[4]) || 0, qty: apply });
        }
        remaining = Math.round((remaining - apply) * 10000) / 10000;
      });
    });
  } catch (e) {
    Logger.log('Erreur _updateStockReservation: ' + e);
  }
  return allDetails;
}

function getResourceReport(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { error: 'Accès refusé.' };
    const reqSheet   = _getSheet('CRAFT_REQUESTS');
    const recipesMap = _buildRecipesMap();
    if (!reqSheet || reqSheet.getLastRow() <= 1) return { byItem: [], total: [] };
    const data = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 19).getValues();
    const accepted = data.filter(function(r) { return r[0] && String(r[8]) === 'Accepté'; });
    const total = {}, byItem = [];
    accepted.forEach(function(r) {
      const cid         = String(r[4]);
      const nomItem     = String(r[5]);
      const qty         = Number(r[6]) || 1;
      const displayName = String(r[3]);
      const customRecipe = _safeParse(r[17], null);
      let resources = [];

      if (customRecipe && Array.isArray(customRecipe)) {
        customRecipe.forEach(function(ing) {
          const used = (Number(ing.qty) || 0) * qty;
          total[ing.resource] = (total[ing.resource] || 0) + used;
          resources.push({ name: ing.resource, qty: used });
        });
      } else if (_isWikiUUID(cid)) {
        _getWikiIngredientsForCraft(cid).forEach(function(ing) {
          const used = ing.qty * qty;
          total[ing.resource] = (total[ing.resource] || 0) + used;
          resources.push({ name: ing.resource, qty: used });
        });
      } else {
        const recipe = recipesMap[cid] || [];
        recipe.forEach(function(ing) {
          const used = ing.qty * qty;
          total[ing.resource] = (total[ing.resource] || 0) + used;
          resources.push({ name: ing.resource, qty: used });
        });
      }
      byItem.push({ displayName: displayName, nomItem: nomItem, quantite: qty, resources: resources,
                    customRecipe: customRecipe ? true : false });
    });
    byItem.sort(function(a, b) { return a.nomItem.localeCompare(b.nomItem); });
    const totalArr = Object.keys(total).map(function(n) { return { name: n, qty: total[n] }; })
                           .sort(function(a, b) { return b.qty - a.qty; });
    return { byItem: byItem, total: totalArr };
  } catch (e) {
    Logger.log('Erreur getResourceReport: ' + e);
    return { byItem: [], total: [] };
  }
}

// ============================================
// STAR CITIZEN WIKI API — BLUEPRINTS SYNC
// WIKI_BLUEPRINTS sheet: ID(0), UUID(1), Name(2), OutputClass(3),
//   CraftTime(4), IngredientsJSON(5), WebURL(6), GameVersion(7), LastUpdated(8), Type(9)
// WIKI_ACTIVE sheet: UUID(0)  — blueprints activés par les admins
// ============================================

var WIKI_BLUEPRINTS_API = 'https://api.star-citizen.wiki/api/blueprints';

function getWikiApiConfig(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { error: 'Accès refusé.' };
    const props = PropertiesService.getScriptProperties();
    return {
      lastSync:     props.getProperty('wiki_last_sync')     || '',
      syncInterval: props.getProperty('wiki_sync_interval') !== null ? parseInt(props.getProperty('wiki_sync_interval')) : null,
      hasTrigger:   ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'syncWikiBlueprints'; })
    };
  } catch(e) { return { lastSync: '', syncInterval: 24, hasTrigger: false }; }
}

function saveWikiApiConfig(config, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    const props = PropertiesService.getScriptProperties();
    if (config.syncInterval !== undefined) props.setProperty('wiki_sync_interval', String(parseInt(config.syncInterval) || 24));
    return { success: true };
  } catch(e) { return { success: false, message: 'Erreur serveur.' }; }
}

function syncWikiBlueprintsAdmin(token) {
  const session = validateToken(token);
  if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
  return _doSyncWiki();
}

// Called by time-based trigger (no token — runs as script owner)
function syncWikiBlueprints() {
  _doSyncWiki();
}

function _doSyncWiki() {
  try {
    const props = PropertiesService.getScriptProperties();
    const allBlueprints = [];
    let page = 1;
    let lastPage = 1;

    // Fetch all pages (100 items/page → ~11 requêtes pour ~1044 blueprints)
    do {
      const url = WIKI_BLUEPRINTS_API + '?page%5Bsize%5D=100&page%5Bnumber%5D=' + page;
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
          'X-Requested-With': 'XMLHttpRequest',
          'User-Agent': 'Mozilla/5.0 (compatible; SibyllaShop/2.3)'
        },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        return { success: false, message: 'Erreur API HTTP ' + response.getResponseCode() + ' (page ' + page + ')' };
      }

      const raw = JSON.parse(response.getContentText());
      const pageData = Array.isArray(raw) ? raw : (raw.data || []);
      if (!Array.isArray(pageData)) return { success: false, message: 'Format API inattendu.' };

      pageData.forEach(function(bp) { allBlueprints.push(bp); });

      if (raw.meta) lastPage = raw.meta.last_page || 1;
      page++;
    } while (page <= lastPage);

    // Fetch fresh detail for active blueprints only (keeps quantities up-to-date without calling 1000+ endpoints)
    const activeUuidsForSync = _getActiveWikiUuids();
    const freshIngredients = {};
    activeUuidsForSync.forEach(function(uuid) {
      try {
        const detailResp = UrlFetchApp.fetch('https://api.star-citizen.wiki/api/blueprints/' + uuid, {
          method: 'GET',
          headers: { 'Accept': 'application/json', 'User-Agent': 'SibyllaShop/2.3' },
          muteHttpExceptions: true
        });
        if (detailResp.getResponseCode() === 200) {
          const bp = JSON.parse(detailResp.getContentText());
          const ingredients = _extractWikiIngredients(bp.data || bp);
          if (ingredients.length > 0) freshIngredients[uuid] = JSON.stringify(ingredients);
        }
      } catch(e) {
        Logger.log('Detail fetch failed for ' + uuid + ': ' + e);
      }
    });

    // Get or create WIKI_BLUEPRINTS sheet
    let sheet = _getSheet('WIKI_BLUEPRINTS');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('WIKI_BLUEPRINTS');
    }

    // Always update header (handles old 9-column sheets)
    sheet.getRange(1, 1, 1, 10).setValues([['ID','UUID','Name','OutputClass','CraftTime','IngredientsJSON','WebURL','GameVersion','LastUpdated','Type']]);
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }

    const now = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
    const rows = [];

    allBlueprints.forEach(function(bp, i) {
      const name = String(bp.output_name || bp.name || '');
      if (!name) return;
      const uuid = String(bp.uuid || '');
      // Use fresh detail data for active blueprints, list API data for others
      const ingredientsJson = freshIngredients[uuid] || JSON.stringify(_extractWikiIngredients(bp));
      rows.push([
        i + 1,
        uuid,
        name,
        String(bp.output_class || (bp.output && bp.output.class) || ''),
        Number(bp.craft_time_seconds || 0),
        ingredientsJson,
        String(bp.web_url || bp.output_item_web_url || ''),
        String(bp.game_version || ''),
        now,
        String((bp.output && bp.output.type) || '')
      ]);
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 10).setValues(rows);
    }

    props.setProperty('wiki_last_sync', now);
    return { success: true, count: rows.length, lastSync: now };
  } catch(e) {
    Logger.log('Erreur _doSyncWiki: ' + e);
    return { success: false, message: String(e) };
  }
}

function _extractWikiIngredients(bp) {
  const result = [];

  // 1. requirement_groups → children (source principale avec quantity_scu)
  if (bp.requirement_groups && bp.requirement_groups.length > 0) {
    bp.requirement_groups.forEach(function(group) {
      _walkWikiChildren(group.children || [], result);
    });
  }

  // 2. tiers[0].requirements si toujours vide
  if (result.length === 0 && bp.tiers && bp.tiers.length > 0) {
    const req = bp.tiers[0].requirements;
    if (req) _walkWikiChildren(req.children || [], result);
  }

  // 3. Fallback: liste plate (pas de quantités)
  if (result.length === 0 && bp.ingredients) {
    bp.ingredients.forEach(function(ing) {
      if (ing.name) result.push({ name: ing.name, qty: 0, minQuality: 0 });
    });
  }

  return result;
}

function _walkWikiChildren(children, result) {
  children.forEach(function(child) {
    const qty = Number(child.quantity_scu || child.quantity || 0);
    if (child.name && qty > 0) {
      result.push({ name: child.name, qty: qty, minQuality: Number(child.min_quality || 0) });
    } else if (child.children && child.children.length > 0) {
      _walkWikiChildren(child.children, result);
    }
  });
}

function getWikiBlueprints() {
  try {
    const sheet = _getSheet('WIKI_BLUEPRINTS');
    if (!sheet || sheet.getLastRow() <= 1) return { blueprints: [], lastSync: '' };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    const blueprints = data.filter(function(r) { return r[0]; }).map(function(r) {
      return {
        id:          Number(r[0]),
        uuid:        String(r[1] || ''),
        name:        String(r[2] || ''),
        outputClass: String(r[3] || ''),
        craftTime:   Number(r[4] || 0),
        ingredients: _safeParse(r[5], []),
        webUrl:      String(r[6] || ''),
        gameVersion: String(r[7] || ''),
        lastUpdated: String(r[8] || ''),
        type:        String(r[9] || '')
      };
    });
    const lastSync = PropertiesService.getScriptProperties().getProperty('wiki_last_sync') || '';
    return { blueprints: blueprints, lastSync: lastSync };
  } catch(e) {
    Logger.log('Erreur getWikiBlueprints: ' + e);
    return { blueprints: [], lastSync: '' };
  }
}

function getWikiBlueprintsAdmin(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { error: 'Accès refusé.' };
    const sheet = _getSheet('WIKI_BLUEPRINTS');
    if (!sheet || sheet.getLastRow() <= 1) return { blueprints: [], lastSync: '' };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    const activeUuids = _getActiveWikiUuids();
    const blueprints = data.filter(function(r) { return r[0]; }).map(function(r) {
      const uuid = String(r[1] || '');
      return {
        id:          Number(r[0]),
        uuid:        uuid,
        name:        String(r[2] || ''),
        outputClass: String(r[3] || ''),
        craftTime:   Number(r[4] || 0),
        ingredients: _safeParse(r[5], []),
        webUrl:      String(r[6] || ''),
        type:        String(r[9] || ''),
        active:      activeUuids.indexOf(uuid) !== -1
      };
    });
    const lastSync = PropertiesService.getScriptProperties().getProperty('wiki_last_sync') || '';
    return { blueprints: blueprints, lastSync: lastSync };
  } catch(e) {
    Logger.log('Erreur getWikiBlueprintsAdmin: ' + e);
    return { blueprints: [], lastSync: '' };
  }
}

function setWikiBlueprintActive(uuid, active, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };

    // Si activation : fetcher le détail pour obtenir les vraies quantités
    if (active) {
      const detailUrl = 'https://api.star-citizen.wiki/api/blueprints/' + String(uuid);
      const resp = UrlFetchApp.fetch(detailUrl, {
        method: 'GET',
        headers: { 'Accept': 'application/json', 'User-Agent': 'SibyllaShop/2.3' },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        const raw = JSON.parse(resp.getContentText());
        const bp = raw.data || raw;
        const ingredients = _extractWikiIngredients(bp);
        if (ingredients.length > 0) {
          // Mettre à jour IngredientsJSON dans WIKI_BLUEPRINTS
          const wikiSheet = _getSheet('WIKI_BLUEPRINTS');
          if (wikiSheet && wikiSheet.getLastRow() > 1) {
            const data = wikiSheet.getRange(2, 1, wikiSheet.getLastRow() - 1, 2).getValues();
            for (var i = 0; i < data.length; i++) {
              if (String(data[i][1]) === String(uuid)) {
                wikiSheet.getRange(i + 2, 6).setValue(JSON.stringify(ingredients));
                break;
              }
            }
          }
        }
      }
    }

    let sheet = _getSheet('WIKI_ACTIVE');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('WIKI_ACTIVE');
      sheet.appendRow(['UUID']);
    }
    const uuids = _getActiveWikiUuids();
    const idx = uuids.indexOf(String(uuid));
    if (active && idx === -1) {
      sheet.appendRow([String(uuid)]);
    } else if (!active && idx !== -1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var j = 0; j < data.length; j++) {
        if (String(data[j][0]) === String(uuid)) { sheet.deleteRow(j + 2); break; }
      }
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: String(e) };
  }
}

function setupWikiTrigger(intervalHours, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'syncWikiBlueprints') ScriptApp.deleteTrigger(t);
    });
    if (parseInt(intervalHours) > 0) {
      ScriptApp.newTrigger('syncWikiBlueprints')
        .timeBased()
        .everyHours(parseInt(intervalHours))
        .create();
    }
    PropertiesService.getScriptProperties().setProperty('wiki_sync_interval', String(intervalHours));
    return { success: true, active: parseInt(intervalHours) > 0 };
  } catch(e) {
    Logger.log('Erreur setupWikiTrigger: ' + e);
    return { success: false, message: String(e) };
  }
}

// ============================================
// LOG EXPORT STOCK (appelé depuis StockAdmin)
// ============================================
function logStockExport(count, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return;
    _logStock(session.displayName || session.username, 'Export CSV : ' + count + ' ligne(s) exportée(s)');
  } catch(e) {}
}

// ============================================
// CRAFTING ENABLE / DISABLE
// ============================================
function getCraftingStatus() {
  try {
    var val = PropertiesService.getScriptProperties().getProperty('crafting_enabled');
    return { enabled: val !== '0' };
  } catch(e) { return { enabled: true }; }
}

function setCraftingEnabled(enabled, token) {
  try {
    var session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    PropertiesService.getScriptProperties().setProperty('crafting_enabled', enabled ? '1' : '0');

    // [DISCORD BOT] Notifier les admins du changement d'état du crafting
    _discordNotifyAdmins(
      (enabled ? '🟢' : '🔴') + ' **Crafting ' + (enabled ? 'activé' : 'désactivé') + '** par ' +
      (session.displayName || session.username)
    );

    return { success: true, enabled: !!enabled };
  } catch(e) { return { success: false, message: String(e) }; }
}

function debugWikiImage() {
  // Fetch item detail pour FS-9 LMG (output_item_uuid du blueprint de test)
  var itemUuid = '6f1674b1-fb58-4661-9114-f418862751d2';
  var url = 'https://api.star-citizen.wiki/api/items/' + itemUuid;
  var resp = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: { 'Accept': 'application/json', 'User-Agent': 'SibyllaShop/2.3' },
    muteHttpExceptions: true
  });
  Logger.log('HTTP: ' + resp.getResponseCode());
  var raw = JSON.parse(resp.getContentText());
  var item = raw.data || raw;
  Logger.log('Clés: ' + Object.keys(item).join(', '));
  Logger.log('thumbnail: ' + JSON.stringify(item.thumbnail));
  Logger.log('media: ' + JSON.stringify(item.media));
  Logger.log('image: ' + JSON.stringify(item.image));
  Logger.log('images: ' + JSON.stringify(item.images));
}

function debugWikiActivation() {
  // UUID du FS-9 LMG (blueprint de test connu)
  var uuid = 'b627e348-29b5-44c8-a836-258df60bcd08';
  var url = 'https://api.star-citizen.wiki/api/blueprints/' + uuid;
  var resp = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: { 'Accept': 'application/json', 'User-Agent': 'SibyllaShop/2.3' },
    muteHttpExceptions: true
  });
  Logger.log('HTTP: ' + resp.getResponseCode());
  var raw = JSON.parse(resp.getContentText());
  Logger.log('Clés racine: ' + Object.keys(raw).join(', '));
  var bp = raw.data || raw;
  Logger.log('output_name: ' + bp.output_name);
  Logger.log('requirement_groups count: ' + (bp.requirement_groups ? bp.requirement_groups.length : 'absent'));
  var ings = _extractWikiIngredients(bp);
  Logger.log('Ingrédients extraits: ' + JSON.stringify(ings));
}

function debugWikiRead() {
  var result = getWikiBlueprints();
  Logger.log('count: ' + (result.blueprints ? result.blueprints.length : 'null'));
  Logger.log('lastSync: ' + result.lastSync);
  if (result.blueprints && result.blueprints.length > 0) {
    Logger.log('premier: ' + JSON.stringify(result.blueprints[0]));
  }
}

function debugWikiApi() {
  var tests = [
    { label: 'Inertia header', url: 'https://api.star-citizen.wiki/blueprints?page%5Bsize%5D=1&page%5Bnumber%5D=1',
      headers: { 'X-Inertia': 'true', 'X-Inertia-Version': '1', 'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0' } },
    { label: 'URL /api/blueprints', url: 'https://api.star-citizen.wiki/api/blueprints?page%5Bsize%5D=1',
      headers: { 'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0' } },
    { label: 'URL /api/v1/blueprints', url: 'https://api.star-citizen.wiki/api/v1/blueprints?page%5Bsize%5D=1',
      headers: { 'Accept': 'application/json', 'User-Agent': 'Mozilla/5.0' } }
  ];
  tests.forEach(function(t) {
    var resp = UrlFetchApp.fetch(t.url, { method:'GET', headers:t.headers, muteHttpExceptions:true });
    var body = resp.getContentText();
    var ct = (resp.getAllHeaders()['Content-Type'] || '');
    Logger.log('[' + t.label + '] HTTP=' + resp.getResponseCode() + ' CT=' + ct + ' body(200c)=' + body.substring(0,200));
  });
}

// ============================================
// CATALOGUE MATÉRIAUX
// MATERIALS_CONFIG sheet: Name(0), Source(1: blueprint|manual), Seuil1(2), Seuil2(3), Seuil3(4)
// Blueprint materials = extraits automatiquement de WIKI_BLUEPRINTS (jamais supprimables)
// Manual materials    = ajoutés manuellement par un admin (supprimables)
// ============================================

function getMaterialsCatalogue(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { error: 'Accès refusé.' };

    // Extraire tous les ingrédients depuis les blueprints
    const bpSheet = _getSheet('WIKI_BLUEPRINTS');
    const bpNames = {};
    if (bpSheet && bpSheet.getLastRow() > 1) {
      bpSheet.getRange(2, 1, bpSheet.getLastRow() - 1, 6).getValues().forEach(function(r) {
        _safeParse(r[5], []).forEach(function(ing) {
          const n = String(ing.name || '').trim();
          if (n) bpNames[n.toLowerCase()] = n;
        });
      });
    }

    // Lire les configs (seuils + matériaux manuels)
    const configSheet = _getSheet('MATERIALS_CONFIG');
    const configs = {};
    if (configSheet && configSheet.getLastRow() > 1) {
      configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 7).getValues().forEach(function(r) {
        if (r[0]) configs[String(r[0]).toLowerCase()] = {
          name: String(r[0]), source: String(r[1] || 'manual'),
          seuil1: Number(r[2]) || 0, seuil2: Number(r[3]) || 0, seuil3: Number(r[4]) || 0,
          cible: Number(r[5]) || 0, categorie: String(r[6] || '')
        };
      });
    }

    const items = [];
    // 1. Blueprint materials (triés)
    Object.keys(bpNames).sort().forEach(function(key) {
      const cfg = configs[key] || {};
      items.push({ name: bpNames[key], source: 'blueprint',
        seuil1: cfg.seuil1 || 0, seuil2: cfg.seuil2 || 0, seuil3: cfg.seuil3 || 0,
        cible: cfg.cible || 0, categorie: cfg.categorie || '' });
    });
    // 2. Matériaux manuels non couverts par les blueprints
    Object.keys(configs).sort().forEach(function(key) {
      if (!bpNames[key] && configs[key].source === 'manual') {
        const cfg = configs[key];
        items.push({ name: cfg.name, source: 'manual',
          seuil1: cfg.seuil1, seuil2: cfg.seuil2, seuil3: cfg.seuil3,
          cible: cfg.cible || 0, categorie: cfg.categorie || '' });
      }
    });

    const lastSync = PropertiesService.getScriptProperties().getProperty('wiki_last_sync') || '';
    return { items: items, lastSync: lastSync };
  } catch(e) {
    Logger.log('Erreur getMaterialsCatalogue: ' + e);
    return { error: String(e) };
  }
}

function saveMaterialSeuils(name, seuil1, seuil2, seuil3, token, cible, categorie) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };

    let sheet = _getSheet('MATERIALS_CONFIG');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('MATERIALS_CONFIG');
      sheet.appendRow(['Name', 'Source', 'Seuil1', 'Seuil2', 'Seuil3', 'Cible', 'Categorie']);
    }
    const nameLow = String(name).toLowerCase();
    if (sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase() === nameLow) {
          sheet.getRange(i + 2, 3, 1, 4).setValues([[Number(seuil1)||0, Number(seuil2)||0, Number(seuil3)||0, Number(cible)||0]]);
          if (categorie !== undefined) sheet.getRange(i + 2, 7).setValue(String(categorie));
          return { success: true };
        }
      }
    }
    // Première config pour ce matériau — on détermine la source
    const bpSheet = _getSheet('WIKI_BLUEPRINTS');
    let source = 'manual';
    if (bpSheet && bpSheet.getLastRow() > 1) {
      const bpData = bpSheet.getRange(2, 1, bpSheet.getLastRow() - 1, 6).getValues();
      outer: for (var j = 0; j < bpData.length; j++) {
        var ings = _safeParse(bpData[j][5], []);
        for (var k = 0; k < ings.length; k++) {
          if (String(ings[k].name || '').toLowerCase() === nameLow) { source = 'blueprint'; break outer; }
        }
      }
    }
    sheet.appendRow([String(name), source, Number(seuil1)||0, Number(seuil2)||0, Number(seuil3)||0, Number(cible)||0, String(categorie||'')]);
    return { success: true };
  } catch(e) { return { success: false, message: String(e) }; }
}

function addManualMaterial(name, token, categorie) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };

    name = String(name).trim();
    if (!name) return { success: false, message: 'Nom requis.' };

    let sheet = _getSheet('MATERIALS_CONFIG');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('MATERIALS_CONFIG');
      sheet.appendRow(['Name', 'Source', 'Seuil1', 'Seuil2', 'Seuil3', 'Cible', 'Categorie']);
    }
    if (sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase() === name.toLowerCase())
          return { success: false, message: 'Ce matériau existe déjà.' };
      }
    }
    sheet.appendRow([name, 'manual', 0, 0, 0, 0, String(categorie||'')]);
    return { success: true };
  } catch(e) { return { success: false, message: String(e) }; }
}

function deleteMaterial(name, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };

    const sheet = _getSheet('MATERIALS_CONFIG');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Matériau introuvable.' };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(name).toLowerCase()) {
        if (String(data[i][1]) !== 'manual')
          return { success: false, message: 'Seuls les matériaux manuels peuvent être supprimés.' };
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, message: 'Matériau introuvable.' };
  } catch(e) { return { success: false, message: String(e) }; }
}

// ============================================
// CATÉGORIES DE STOCK
// STOCK_CATEGORIES sheet: Name(0), UnitType(1)   UnitType: 'cSCU' | 'unité'
// ============================================

// Fonction de debug — à lancer depuis l'éditeur GAS (Exécuter > debugCraftDolivine)
function debugCraftDolivine() {
  var catUnitMap   = _getCategoryUnitMap();
  var unitFactors  = _getResourceUnitFactors();
  var stockMap     = _buildStockMap();

  Logger.log('=== STOCK_CATEGORIES (catUnitMap) ===');
  Logger.log(JSON.stringify(catUnitMap));

  Logger.log('=== MATERIALS_CONFIG pour Dolivine ===');
  var cfgSheet = _getSheet('MATERIALS_CONFIG');
  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, 7).getValues().forEach(function(r) {
      if (String(r[0]).toLowerCase() === 'dolivine') Logger.log('Trouvé: ' + JSON.stringify(r));
    });
  }

  Logger.log('=== STOCK rows pour Dolivine ===');
  var stockSheet = _getSheet('STOCK');
  if (stockSheet && stockSheet.getLastRow() > 1) {
    stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 5).getValues().forEach(function(r) {
      if (String(r[2]).toLowerCase() === 'dolivine') Logger.log('Trouvé: ' + JSON.stringify(r));
    });
  }

  var factor = unitFactors['dolivine'] !== undefined ? unitFactors['dolivine'] : 100;
  Logger.log('=== RÉSULTAT ===');
  Logger.log('unitFactors["dolivine"] = ' + unitFactors['dolivine'] + '  (undefined = fallback 100)');
  Logger.log('stockMap["dolivine"]    = ' + stockMap['dolivine']);
  Logger.log('Quantité requise        = 4 × ' + factor + ' = ' + (4 * factor));
  Logger.log('Craftable               = ' + ((stockMap['dolivine'] || 0) >= 4 * factor));
}

function _getCategoryUnitMap() {
  const map = {};
  const sheet = _getSheet('STOCK_CATEGORIES');
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function(r) {
      if (r[0]) map[String(r[0]).toLowerCase()] = String(r[1] || 'cSCU');
    });
  }
  return map;
}

function getCategories(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { error: 'Accès refusé.' };
    const sheet = _getSheet('STOCK_CATEGORIES');
    const cats = [];
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function(r) {
        if (r[0]) cats.push({ name: String(r[0]), unitType: String(r[1] || 'cSCU') });
      });
    }
    cats.sort(function(a,b){ return a.name.localeCompare(b.name,'fr'); });
    return { categories: cats };
  } catch(e) { return { error: String(e) }; }
}

function saveCategory(name, unitType, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    name = String(name).trim();
    if (!name) return { success: false, message: 'Nom requis.' };
    unitType = (unitType === 'unité') ? 'unité' : 'cSCU';
    let sheet = _getSheet('STOCK_CATEGORIES');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('STOCK_CATEGORIES');
      sheet.appendRow(['Name', 'UnitType']);
    }
    if (sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase() === name.toLowerCase()) {
          sheet.getRange(i + 2, 2).setValue(unitType);
          return { success: true };
        }
      }
    }
    sheet.appendRow([name, unitType]);
    return { success: true };
  } catch(e) { return { success: false, message: String(e) }; }
}

function deleteCategory(name, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    const sheet = _getSheet('STOCK_CATEGORIES');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Catégorie introuvable.' };
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(name).toLowerCase()) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, message: 'Catégorie introuvable.' };
  } catch(e) { return { success: false, message: String(e) }; }
}

// ============================================
// IMPORT STOCK CSV
// rows: [{categorie, item, qualite, quantite, seuil1, seuil2, seuil3, actif, imageUrl}]
// - Upsert stock entries (match par item + qualite)
// - Ajoute au catalogue (MATERIALS_CONFIG) les items absents
// ============================================

function importStockFromCsv(rows, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    if (!Array.isArray(rows) || rows.length === 0) return { success: false, message: 'Aucune donnée.' };

    const stockSheet = _getSheet('STOCK');
    if (!stockSheet) return { success: false, message: 'Sheet STOCK introuvable.' };

    // Charger le stock existant (ID(0), Item(1)…qualite(4)…)
    const stockData = stockSheet.getLastRow() > 1
      ? stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, stockSheet.getLastColumn()).getValues()
      : [];

    // Trouver les headers (ligne 1) pour connaître les colonnes
    const headers = stockSheet.getRange(1, 1, 1, stockSheet.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).toLowerCase(); });
    const col = function(name) { return headers.indexOf(name); };
    const iItem    = col('item') !== -1 ? col('item') : 1;
    const iQualite = col('qualite') !== -1 ? col('qualite') : 4;
    const iQty     = col('quantite') !== -1 ? col('quantite') : 5;
    const iCat     = col('categorie') !== -1 ? col('categorie') : 2;
    const iImg     = col('imageurl') !== -1 ? col('imageurl') : -1;
    const iS1      = col('seuil1') !== -1 ? col('seuil1') : -1;
    const iS2      = col('seuil2') !== -1 ? col('seuil2') : -1;
    const iS3      = col('seuil3') !== -1 ? col('seuil3') : -1;
    const iActif   = col('actif') !== -1 ? col('actif') : -1;

    // Index du stock existant : "item::qualite" → row index (0-based in stockData)
    const existingIndex = {};
    stockData.forEach(function(r, i) {
      const key = String(r[iItem]).toLowerCase() + '::' + String(r[iQualite]);
      existingIndex[key] = i;
    });

    // Collecter les noms existants dans le catalogue
    const bpSheet = _getSheet('WIKI_BLUEPRINTS');
    const bpNames = {};
    if (bpSheet && bpSheet.getLastRow() > 1) {
      bpSheet.getRange(2, 1, bpSheet.getLastRow() - 1, 6).getValues().forEach(function(r) {
        _safeParse(r[5], []).forEach(function(ing) {
          const n = String(ing.name || '').trim().toLowerCase();
          if (n) bpNames[n] = true;
        });
      });
    }
    let configSheet = _getSheet('MATERIALS_CONFIG');
    if (!configSheet) {
      configSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('MATERIALS_CONFIG');
      configSheet.appendRow(['Name','Source','Seuil1','Seuil2','Seuil3']);
    }
    const configNames = {};
    if (configSheet.getLastRow() > 1) {
      configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
        if (r[0]) configNames[String(r[0]).toLowerCase()] = true;
      });
    }

    let added = 0, updated = 0, newCatalogue = 0;
    const now = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');

    rows.forEach(function(row) {
      if (!row.item) return;
      const key = row.item.toLowerCase() + '::' + (row.qualite || 500);
      const existIdx = existingIndex[key];

      if (existIdx !== undefined) {
        // Mise à jour
        const sheetRow = existIdx + 2;
        if (iQty >= 0)   stockSheet.getRange(sheetRow, iQty + 1).setValue(Number(row.quantite) || 0);
        if (iCat >= 0)   stockSheet.getRange(sheetRow, iCat + 1).setValue(row.categorie || 'Matériaux');
        if (iS1 >= 0)    stockSheet.getRange(sheetRow, iS1 + 1).setValue(Number(row.seuil1) || 0);
        if (iS2 >= 0)    stockSheet.getRange(sheetRow, iS2 + 1).setValue(Number(row.seuil2) || 0);
        if (iS3 >= 0)    stockSheet.getRange(sheetRow, iS3 + 1).setValue(Number(row.seuil3) || 0);
        if (iActif >= 0) stockSheet.getRange(sheetRow, iActif + 1).setValue(row.actif !== false);
        if (iImg >= 0 && row.imageUrl) stockSheet.getRange(sheetRow, iImg + 1).setValue(row.imageUrl);
        updated++;
      } else {
        // Ajout via addStockItem (réutilise la logique existante)
        addStockItem({
          categorie: row.categorie || 'Matériaux',
          item:      row.item,
          qualite:   Number(row.qualite) || 500,
          quantite:  Number(row.quantite) || 0,
          unite:     'SCU',
          imageUrl:  row.imageUrl || '',
          seuil1:    Number(row.seuil1) || 0,
          seuil2:    Number(row.seuil2) || 0,
          seuil3:    Number(row.seuil3) || 0,
          actif:     row.actif !== false
        }, token);
        added++;
      }

      // Ajouter au catalogue si absent
      const nameLow = row.item.toLowerCase();
      if (!bpNames[nameLow] && !configNames[nameLow]) {
        configSheet.appendRow([row.item, 'manual', Number(row.seuil1)||0, Number(row.seuil2)||0, Number(row.seuil3)||0]);
        configNames[nameLow] = true;
        newCatalogue++;
      }
    });

    _logStock(session.displayName || session.username,
      'Import CSV : ' + added + ' ajouté(s), ' + updated + ' mis à jour, ' + newCatalogue + ' nouveau(x) au catalogue');
    return { success: true, updated: updated, added: added, newCatalogue: newCatalogue };
  } catch(e) {
    Logger.log('Erreur importStockFromCsv: ' + e);
    return { success: false, message: String(e) };
  }
}

// ============================================
// ENDPOINT WEB APP
// ============================================

function doGet(e) {
  const page   = (e && e.parameter && e.parameter.page)  || 'index';
  const token  = (e && e.parameter && e.parameter.token) || '';
  const from   = (e && e.parameter && e.parameter.from)  || '';
  const appUrl = ScriptApp.getService().getUrl();
  let currentUserRole = 0;

  // Pages protégées membres (connexion simple requise)
  const memberPages = ['stock', 'craft'];
  if (memberPages.indexOf(page) !== -1) {
    const session = validateToken(token);
    if (!session) {
      const tpl = HtmlService.createTemplateFromFile('Login');
      tpl.appUrl   = appUrl;
      tpl.errorMsg = '';
      tpl.fromPage = page;
      return tpl.evaluate().setTitle('Connexion - Sibylla');
    }
    currentUserRole = session.role || 0;
  }

  // Pages protégées admin
  const adminPages = ['vendeur', 'stock-admin', 'craft-admin', 'stock-chart'];
  if (adminPages.indexOf(page) !== -1) {
    const session = validateToken(token);
    if (!session) {
      const tpl = HtmlService.createTemplateFromFile('Login');
      tpl.appUrl   = appUrl;
      tpl.errorMsg = 'Session expirée. Reconnectez-vous.';
      tpl.fromPage = page;
      return tpl.evaluate().setTitle('Connexion - Sibylla');
    }
    if ((page === 'stock-admin' || page === 'craft-admin' || page === 'stock-chart') && (session.role || 0) < 1) {
      const tpl = HtmlService.createTemplateFromFile('Login');
      tpl.appUrl   = appUrl;
      tpl.errorMsg = 'Accès réservé aux gestionnaires.';
      tpl.fromPage = page;
      return tpl.evaluate().setTitle('Connexion - Sibylla');
    }
    currentUserRole = session.role || 0;
  }

  const pages = {
    'index':       { file: 'Index',        title: 'Sibylla Shop' },
    'market':      { file: 'Marketplace',  title: 'Market - Sibylla' },
    'marketplace': { file: 'Marketplace',  title: 'Market - Sibylla' },
    'login':       { file: 'Login',        title: 'Connexion - Sibylla' },
    'vendeur':     { file: 'VendeurPanel', title: 'Gestion Vendeur - Sibylla' },
    'stock':       { file: 'Stock',        title: 'Stock - Sibylla' },
    'craft':       { file: 'Craft',        title: 'Blueprint - Sibylla' },
    'stock-admin': { file: 'StockAdmin',   title: 'Stock Admin - Sibylla' },
    'craft-admin': { file: 'CraftAdmin',   title: 'Blueprint Admin - Sibylla' },
    'stock-chart': { file: 'StockChart',   title: 'Graphiques Stock - Sibylla' }
  };

  const p = pages[page] || pages['index'];
  const tpl = HtmlService.createTemplateFromFile(p.file);
  tpl.appUrl        = appUrl;
  tpl.vendeurToken  = token;
  tpl.userRole      = currentUserRole;
  tpl.errorMsg      = '';
  tpl.fromPage      = from;
  tpl.resourceParam = (e && e.parameter && e.parameter.resource) || '';
  return tpl.evaluate().setTitle(p.title);
}


// ============================================================
// ===== DISCORD — NOTIFICATIONS VIA WEBHOOK =====
// ============================================================
//
// MISE EN PLACE (une seule fois) :
//
//   1. Dans ton serveur Discord, crée un salon texte dédié, ex: #sibylla-notifs
//   2. Paramètres du salon → Intégrations → Webhooks → Créer un webhook
//   3. Copie l'URL du webhook
//   4. Dans l'éditeur Apps Script → ⚙️ Paramètres du projet → Propriétés de script, ajoute :
//        Clé   : DISCORD_WEBHOOK_URL
//        Valeur: l'URL copiée (commence par https://discord.com/api/webhooks/...)
//
//   Les messages apparaîtront dans le salon avec mention (@) des personnes concernées.
//   La colonne F de la feuille USERS doit contenir le Discord User ID de chaque membre
//   (le numéro — visible en mode développeur Discord : clic droit sur pseudo → Copier l'identifiant)
//
// ============================================================

/**
 * Récupère le Discord User ID d'un username depuis la feuille USERS (colonne F).
 * Retourne une chaîne vide si non trouvé.
 */
function _getDiscordId(username) {
  try {
    const sheet = _getSheet('USERS');
    if (!sheet) return '';
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(username)) {
        return String(data[i][5] || '');
      }
    }
    return '';
  } catch (e) {
    Logger.log('[Discord] _getDiscordId erreur: ' + e);
    return '';
  }
}

/**
 * Ajoute un message à la file d'attente Discord (Script Properties).
 * Le trigger flushDiscordQueue (toutes les minutes) envoie les messages un par un.
 * Cela évite le rate-limit Discord (429) et garantit la livraison séquentielle.
 */
function _discordWebhook(message, discordIds) {
  try {
    const url = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL');
    if (!url) {
      Logger.log('[Discord] DISCORD_WEBHOOK_URL non configuré — notification non envoyée.');
      return;
    }
    const props = PropertiesService.getScriptProperties();
    const queue = _safeParse(props.getProperty('DISCORD_QUEUE'), []);
    queue.push({ message: String(message), ids: discordIds || [], ts: Date.now() });
    // Limiter la queue à 100 messages max
    while (queue.length > 100) queue.shift();
    props.setProperty('DISCORD_QUEUE', JSON.stringify(queue));
  } catch (e) {
    Logger.log('[Discord] _discordWebhook (enqueue) erreur: ' + e);
  }
}

/**
 * Envoie UN message de la file Discord. Appelé par un trigger toutes les minutes.
 * Appeler setupDiscordQueueTrigger() une fois pour activer.
 */
function flushDiscordQueue() {
  try {
    const props = PropertiesService.getScriptProperties();
    const url = props.getProperty('DISCORD_WEBHOOK_URL');
    if (!url) return;
    const queue = _safeParse(props.getProperty('DISCORD_QUEUE'), []);
    if (queue.length === 0) return;
    const item = queue.shift();
    // Toujours sauvegarder la queue réduite avant l'envoi (évite doublons si l'envoi crashe)
    props.setProperty('DISCORD_QUEUE', JSON.stringify(queue));
    let mentions = '';
    if (item.ids && item.ids.length > 0) {
      mentions = item.ids.filter(Boolean).map(function(id) { return '<@' + id + '>'; }).join(' ') + '\n';
    }
    const payload = {
      content: (mentions + String(item.message)).substring(0, 2000),
      username: 'Sibylla'
    };
    const resp = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() === 429) {
      // Rate limited : remettre le message en tête de file
      queue.unshift(item);
      props.setProperty('DISCORD_QUEUE', JSON.stringify(queue));
      Logger.log('[Discord] Rate-limit 429, message remis en file.');
    } else if (resp.getResponseCode() !== 204 && resp.getResponseCode() !== 200) {
      Logger.log('[Discord] Erreur envoi — HTTP ' + resp.getResponseCode() + ' : ' + resp.getContentText());
    }
  } catch (e) {
    Logger.log('[Discord] flushDiscordQueue erreur: ' + e);
  }
}

/**
 * Active le trigger toutes les minutes pour flushDiscordQueue.
 * À exécuter une seule fois depuis l'éditeur Apps Script.
 */
function setupDiscordQueueTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'flushDiscordQueue') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('flushDiscordQueue').timeBased().everyMinutes(1).create();
  Logger.log('[Discord] Trigger flushDiscordQueue activé (toutes les minutes).');
}

/**
 * Notifie un utilisateur Sibylla dans le salon webhook (avec mention @).
 */
function _discordNotifyUser(username, message) {
  const discordId = _getDiscordId(username);
  _discordWebhook(message, discordId ? [discordId] : []);
}

/**
 * Notifie tous les admins (role >= 1) dans le salon webhook (avec mentions @).
 */
function _discordNotifyAdmins(message) {
  try {
    const sheet = _getSheet('USERS');
    if (!sheet) { _discordWebhook(message, []); return; }
    const data = sheet.getDataRange().getValues();
    const adminIds = [];
    for (let i = 1; i < data.length; i++) {
      if ((Number(data[i][4]) || 0) >= 1 && data[i][5]) {
        adminIds.push(String(data[i][5]));
      }
    }
    _discordWebhook(message, adminIds);
  } catch (e) {
    Logger.log('[Discord] _discordNotifyAdmins erreur: ' + e);
  }
}

// ============================================================
// ===== DISCORD — FONCTION DE TEST =====
// Sélectionner "testDiscordWebhook" dans l'éditeur Apps Script puis cliquer ▶
// ============================================================
function testDiscordWebhook() {
  Logger.log('[Test Discord] Envoi d\'un message test via webhook...');
  _discordWebhook('✅ **Test Sibylla**\nLe webhook fonctionne correctement !', []);
  Logger.log('[Test Discord] Terminé. Vérifie le salon Discord.');
}

// ============================================
// OBJECTIFS DE STOCK — cible par item
// Stockée dans MATERIALS_CONFIG colonne 6 (Cible)
// ============================================

function setStockTarget(name, cible, token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    let sheet = _getSheet('MATERIALS_CONFIG');
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('MATERIALS_CONFIG');
      sheet.appendRow(['Name', 'Source', 'Seuil1', 'Seuil2', 'Seuil3', 'Cible']);
    }
    const nameLow = String(name).toLowerCase();
    if (sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase() === nameLow) {
          sheet.getRange(i + 2, 6).setValue(Number(cible) || 0);
          return { success: true };
        }
      }
    }
    // Item absent du catalogue : l'ajouter comme entrée manuelle
    const bpSheet = _getSheet('WIKI_BLUEPRINTS');
    let source = 'manual';
    if (bpSheet && bpSheet.getLastRow() > 1) {
      const bpData = bpSheet.getRange(2, 1, bpSheet.getLastRow() - 1, 6).getValues();
      outer: for (let j = 0; j < bpData.length; j++) {
        const ings = _safeParse(bpData[j][5], []);
        for (let k = 0; k < ings.length; k++) {
          if (String(ings[k].name || '').toLowerCase() === nameLow) { source = 'blueprint'; break outer; }
        }
      }
    }
    sheet.appendRow([name, source, 0, 0, 0, Number(cible) || 0]);
    return { success: true };
  } catch(e) { return { success: false, message: String(e) }; }
}

// ============================================
// HISTORIQUE STOCK — snapshots journaliers
// STOCK_HISTORY: Date(0), ItemName(1), TotalAvail(2)
// Appeler setupStockHistoryTrigger() une fois pour activer le snapshot quotidien.
// ============================================

function snapshotStockHistory() {
  try {
    const sheet = _getSheet('STOCK');
    if (!sheet || sheet.getLastRow() <= 1) return;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
    const totals = {};
    data.forEach(function(r) {
      if (!r[0] || !_isActif(r[11])) return;
      const name = String(r[2]);
      const avail = Math.max(0, (Number(r[3]) || 0) - (Number(r[7]) || 0));
      totals[name] = (totals[name] || 0) + avail;
    });
    let histSheet = _getSheet('STOCK_HISTORY');
    if (!histSheet) {
      histSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('STOCK_HISTORY');
      histSheet.appendRow(['Date', 'ItemName', 'TotalAvail']);
    }
    const date = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');
    const rows = Object.keys(totals).map(function(name) { return [date, name, totals[name]]; });
    if (rows.length > 0) histSheet.getRange(histSheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    _pruneStockHistory(90);
  } catch(e) { Logger.log('Erreur snapshotStockHistory: ' + e); }
}

function _pruneStockHistory(keepDays) {
  try {
    const sheet = _getSheet('STOCK_HISTORY');
    if (!sheet || sheet.getLastRow() <= 1) return;
    const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - keepDays);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const parts = _fmtDay(data[i][0]).split('/');
      const d = new Date(parts[2] + '-' + parts[1] + '-' + parts[0]);
      if (!isNaN(d.getTime()) && d < cutoff) sheet.deleteRow(i + 2);
    }
  } catch(e) {}
}

function setupStockHistoryTrigger(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'snapshotStockHistory') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('snapshotStockHistory').timeBased().everyDays(1).atHour(3).create();
    Logger.log('[History] Trigger snapshotStockHistory activé (chaque jour à 3h).');
    return { success: true };
  } catch(e) { return { success: false, message: String(e) }; }
}

function getStockHistory(itemNames, days, token) {
  try {
    const session = validateToken(token);
    if (!session) return { error: 'Connexion requise.' };
    const sheet = _getSheet('STOCK_HISTORY');
    if (!sheet || sheet.getLastRow() <= 1) return { history: {} };
    const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - (days || 14));
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    const names = Array.isArray(itemNames)
      ? itemNames.map(function(n) { return String(n).toLowerCase(); })
      : [];
    const result = {};
    data.forEach(function(r) {
      if (!r[0] || !r[1]) return;
      const name = String(r[1]);
      if (names.length > 0 && names.indexOf(name.toLowerCase()) === -1) return;
      const dateStr = _fmtDay(r[0]);
      const parts = dateStr.split('/');
      if (parts.length !== 3) return;
      const d = new Date(parts[2] + '-' + parts[1] + '-' + parts[0]);
      if (!isNaN(d.getTime()) && d >= cutoff) {
        if (!result[name]) result[name] = [];
        result[name].push({ date: dateStr, val: Number(r[2]) || 0 });
      }
    });
    return { history: result };
  } catch(e) { Logger.log('Erreur getStockHistory: ' + e); return { history: {} }; }
}

function backfillStockHistory(token) {
  try {
    const session = validateToken(token);
    if (!session || (session.role || 0) < 1) return { success: false, message: 'Accès refusé.' };

    const logSheet = _getSheet('STOCK_LOG');
    if (!logSheet || logSheet.getLastRow() <= 1) return { success: false, message: 'Aucune entrée dans STOCK_LOG.' };

    const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getValues();

    function parseRawDate(raw) {
      if (raw instanceof Date) return raw;
      const s = String(raw);
      const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) return new Date(m[3] + '-' + m[2] + '-' + m[1] + 'T12:00:00');
      return new Date(s);
    }
    function toDayStr(d) {
      return ('0'+d.getDate()).slice(-2) + '/' + ('0'+(d.getMonth()+1)).slice(-2) + '/' + d.getFullYear();
    }

    // Parse log into structured entries
    const entries = [];
    logData.forEach(function(r) {
      if (!r[0]) return;
      const d = parseRawDate(r[1]);
      if (isNaN(d.getTime())) return;
      const details = String(r[3]);
      let action, item, qty, quality;
      let m;

      m = details.match(/^Ajout\s*:\s*(.+?)\s+x(\d+)\s+\(Q:(\d+)\)/);
      if (m) { action = 'add'; item = m[1].trim(); qty = parseInt(m[2]); quality = m[3]; }

      if (!action) {
        m = details.match(/^Modif\s*:\s*(.+?)\s*(?:→|->)\s*x(\d+)\s+\(Q:(\d+)\)/);
        if (m) { action = 'set'; item = m[1].trim(); qty = parseInt(m[2]); quality = m[3]; }
      }
      if (!action) {
        m = details.match(/^Suppression\s*:\s*(.+?)\s+x(\d+)\s+\(Q:(\d+)\)/);
        if (m) { action = 'del'; item = m[1].trim(); qty = parseInt(m[2]); quality = m[3]; }
      }
      if (!action) return;

      entries.push({ dayStr: toDayStr(d), ts: d.getTime(), action: action, item: item, qty: qty, key: item + '|' + quality });
    });

    if (entries.length === 0) return { success: false, message: 'Aucune entrée parseable dans le log.' };

    entries.sort(function(a, b) { return a.ts - b.ts; });

    // Group by day
    const dayMap = {};
    entries.forEach(function(e) {
      if (!dayMap[e.dayStr]) dayMap[e.dayStr] = [];
      dayMap[e.dayStr].push(e);
    });

    // Iterate day by day from first to today, carry forward state
    function strToDate(s) {
      const p = s.split('/');
      return new Date(p[2] + '-' + p[1] + '-' + p[0] + 'T12:00:00');
    }

    const state = {}; // "Item|Quality" → qty
    const today = new Date(); today.setHours(23, 59, 59, 0);
    const cursor = strToDate(entries[0].dayStr);
    const allRows = [];

    while (cursor <= today) {
      const ds = toDayStr(cursor);

      if (dayMap[ds]) {
        dayMap[ds].forEach(function(e) {
          if (e.action === 'add') {
            state[e.key] = (state[e.key] || 0) + e.qty;
          } else if (e.action === 'set') {
            state[e.key] = e.qty;
          } else if (e.action === 'del') {
            const remaining = Math.max(0, (state[e.key] || 0) - e.qty);
            if (remaining === 0) delete state[e.key];
            else state[e.key] = remaining;
          }
        });
      }

      // Compute per-item totals and write one row per item
      const totals = {};
      Object.keys(state).forEach(function(k) {
        const itemName = k.split('|')[0];
        totals[itemName] = (totals[itemName] || 0) + (state[k] || 0);
      });
      Object.keys(totals).forEach(function(itemName) {
        allRows.push([ds, itemName, totals[itemName]]);
      });

      cursor.setDate(cursor.getDate() + 1);
    }

    if (allRows.length === 0) return { success: false, message: 'Aucune donnée calculée.' };

    // Write to STOCK_HISTORY (clear existing, rewrite all)
    let histSheet = _getSheet('STOCK_HISTORY');
    if (!histSheet) {
      histSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('STOCK_HISTORY');
      histSheet.appendRow(['Date', 'ItemName', 'TotalAvail']);
    } else {
      if (histSheet.getLastRow() > 1) histSheet.deleteRows(2, histSheet.getLastRow() - 1);
    }
    histSheet.getRange(2, 1, allRows.length, 3).setValues(allRows);

    return { success: true, count: allRows.length, items: Object.keys(allRows.reduce(function(acc, r){ acc[r[1]]=1; return acc; }, {})).length };
  } catch(e) { Logger.log('Erreur backfillStockHistory: ' + e); return { success: false, message: String(e) }; }
}
